# LINQ for ShapeSheet queries — before / after

Motivation doc for [#170](https://github.com/saveenr/VisioAutomation/issues/170). Lays out a list of common ShapeSheet-query scenarios and shows each one twice: how it reads in the codebase today, and how it might read with a LINQ-shaped surface. The "after" code is *illustrative*: it shows the surface a user would see, not a claim about what the spike will land on.

The spike that follows ([#170](https://github.com/saveenr/VisioAutomation/issues/170)) decides whether any of these "after" forms is worth pursuing, and if so which option (lightweight `IEnumerable<T>` extensions, full `IQueryable<T>` provider, or a hybrid). This doc deliberately stays one level above that decision: its job is to make the *user-side* benefit concrete enough that the spike has a target to aim at.

## Why LINQ might help

The existing `CellQuery` / `SectionQuery` builder API works and is already the engine behind every typed `*Cells.cs` record (`ShapeXFormCells`, `ShapeFormatCells`, `CustomPropertyCells`, ...). Five recurring frictions show up at the call sites:

1. **Two-step indirection at the read site.** You add a column to get back a `col_fg` index, then index `formulas[0][col_fg]` to get the value. The column variable is dead between those two lines; its only job is to bridge the build-then-index gap.
2. **Cross-category cells force a drop to raw `CellQuery`.** `ShapeXFormCells` has `PinX` / `Width`; `ShapeFormatCells` has `FillForeground` / `LineWeight`. Want both at once? Today you build a `CellQuery` by hand and lose the typed names.
3. **Ad-hoc projection over-fetches.** Want just two cells out of `ShapeFormatCells`? `ShapeFormatCells.GetCells` always pays for all 25.
4. **Filter / order / group happen *after* the query, not as part of it.** "Shapes wider than 2 inches" means: query every shape, materialize, filter in C#. The query shape doesn't compose with predicates.
5. **Variable-row sections produce deeply-indexed shapes.** `out_formulas[shape_idx][section_idx][row_idx][col_idx]` — four levels of indexing, no names. See [`SectionRowHandling`](../../VisioAutomation_2010/VTest/Core/ShapeSheet/ShapeSheetQueryTests.cs#L127).

LINQ is the standard C# answer to (1), (3), (4), and (5). It's at best partial relief for (2) — naming a mixed-category record still requires *some* declaration, even if the declaration is a `new { ... }`. And it may be *no* help on the variable-row case if the projection shape can't naturally express "n rows per shape, m columns per row, both n and m varying."

The scenarios below try to be honest about which of those buckets each case falls into.

## Conventions used in the "after" code

- `page.Shapes.Query()` — hypothetical entry point that returns an `IQueryable<ShapeSheetRow>` (or equivalent) over the shapes on a page. Concrete shape TBD by the spike.
- `s.PinX`, `s.FillForeground`, etc. — typed cell accessors mirroring the existing `*Cells.cs` property names. Source-of-truth for naming is the existing records, so `XFormPinX` stays `PinX`, `CustomPropValue` stays `Formula` (per the [#144](https://github.com/saveenr/VisioAutomation/issues/144) rename in `CustomPropertyCells`), etc.
- `.Formulas()` / `.Results<T>()` — terminal operators that close the query and force the COM call. The lazy intermediate stages (`Where`, `Select`, `OrderBy`) accumulate the column set; the terminal operator builds and runs the underlying `CellQuery` in one batch.

The intent is that the *user* never builds a `CellQuery` by hand; the LINQ surface translates their projection / predicate / order into one batched call.

## Scenarios

Each scenario links to the existing test it's adapted from. The "today" snippets are trimmed to the query bit — setup (page creation, shape drawing, formula seeding) is elided when not load-bearing for the comparison.

### 1. Read one typed group of cells from one shape

**Source:** [`ResultsInt_SingleShape`](../../VisioAutomation_2010/VTest/Core/ShapeSheet/ShapeSheetWriterTests.cs#L56), [`ShapeFormatCells.GetCells`](../../VisioAutomation_2010/VisioAutomation/Shapes/ShapeFormatCells.cs).

This is the case where today's API is *already good*. Including it for honesty: not every "before" needs an "after."

**Today:**

```csharp
var fmt = ShapeFormatCells.GetCells(shape, CellValueType.Formula);
string fg = fmt.FillForeground;
string bg = fmt.FillBackground;
string pat = fmt.FillPattern;
```

**With LINQ:**

```csharp
var fmt = shape.Query<ShapeFormatCells>(CellValueType.Formula);
string fg = fmt.FillForeground;
string bg = fmt.FillBackground;
string pat = fmt.FillPattern;
```

**What this shows:** Roughly a wash. The existing typed-record API already gives names and a single-call shape; LINQ here is a renamed entry point with no real ergonomic gain. **A LINQ surface that *only* delivers on this case isn't worth shipping.**

### 2. Project an ad-hoc subset from one shape

**Source:** Adapted from [`GetResults_SingleShape`](../../VisioAutomation_2010/VTest/Core/ShapeSheet/ShapeSheetQueryTests.cs#L15) and [`ResultsInt_SingleShape`](../../VisioAutomation_2010/VTest/Core/ShapeSheet/ShapeSheetWriterTests.cs#L56). Want just the fill triple, not the full 25-cell `ShapeFormatCells`.

**Today (option A, typed record, over-fetches):**

```csharp
var fmt = ShapeFormatCells.GetCells(shape, CellValueType.Formula);
string fg = fmt.FillForeground;
string bg = fmt.FillBackground;
string pat = fmt.FillPattern;
// 22 other cells were also fetched and discarded.
```

**Today (option B, raw CellQuery, no over-fetch but loses names):**

```csharp
var query = new VASS.Query.CellQuery();
var col_fg  = query.Columns.Add(VA.Core.SrcConstants.FillForeground);
var col_bg  = query.Columns.Add(VA.Core.SrcConstants.FillBackground);
var col_pat = query.Columns.Add(VA.Core.SrcConstants.FillPattern);

var formulas = query.GetFormulas(shape);
string fg  = formulas[0][col_fg];
string bg  = formulas[0][col_bg];
string pat = formulas[0][col_pat];
```

**With LINQ:**

```csharp
var fill = shape.Query(s => new { s.FillForeground, s.FillBackground, s.FillPattern })
                .Formulas();
// fill.FillForeground, fill.FillBackground, fill.FillPattern
```

**What this shows:** First case where LINQ pulls weight. The projection drives the column set, so only three cells are fetched, the names are kept, and the build-then-index ceremony from option B is gone. Today's user has to choose between fetching too much (A) or losing names (B); the projection collapses the choice.

### 3. Mix cells across categories

**Source:** Synthetic. Want PinX + Width from `ShapeXFormCells` *and* FillForeground from `ShapeFormatCells` in one call.

**Today:**

```csharp
var query = new VASS.Query.CellQuery();
var col_pinx  = query.Columns.Add(VA.Core.SrcConstants.XFormPinX);
var col_width = query.Columns.Add(VA.Core.SrcConstants.XFormWidth);
var col_fg    = query.Columns.Add(VA.Core.SrcConstants.FillForeground);

var formulas = query.GetFormulas(shape);
double pinx   = double.Parse(formulas[0][col_pinx]);   // or use GetResults<double>
double width  = double.Parse(formulas[0][col_width]);
string fg     = formulas[0][col_fg];
```

(Or two separate calls — `ShapeXFormCells.GetCells` plus `ShapeFormatCells.GetCells` — which doubles the round-trip.)

**With LINQ:**

```csharp
var info = shape.Query(s => new { s.PinX, s.Width, s.FillForeground })
                .Results<string>();
// info.PinX, info.Width, info.FillForeground — one batched call
```

**What this shows:** The cross-category case is exactly where the typed-record API today falls off a cliff and forces a manual `CellQuery`. LINQ erases the boundary: the projector is just a record literal that happens to mention cells from three different categories.

### 4. Read cells from many shapes (multi-shape batching)

**Source:** [`GetFormulasAndResults_OnThreeShapes_ReturnsValuesPerShapePerColumn`](../../VisioAutomation_2010/VTest/Core/ShapeSheet/ShapeSheetQueryTests.cs#L188).

**Today:**

```csharp
var shapeids = new List<int> { shape_a.ID, shape_b.ID, shape_c.ID };

var query = new VASS.Query.CellQuery();
var col_pinx = query.Columns.Add(VA.Core.SrcConstants.XFormPinX);
var col_piny = query.Columns.Add(VA.Core.SrcConstants.XFormPinY);

var results = query.GetResults<double>(page1, shapeids);
// results[0][col_pinx], results[0][col_piny], ..., results[2][col_piny]
```

**With LINQ:**

```csharp
var positions = page1.Shapes
    .Where(s => new[] { shape_a.ID, shape_b.ID, shape_c.ID }.Contains(s.ID))
    .Query(s => new { s.ID, s.PinX, s.PinY })
    .Results<double>()
    .ToList();

foreach (var p in positions)
    Console.WriteLine($"{p.ID}: ({p.PinX}, {p.PinY})");
```

**What this shows:** Two wins. (1) The result is a flat sequence of named records keyed by shape, not a 2D `[shape_idx][col_idx]` table — `foreach` works naturally and the column-index variables disappear. (2) The shape-selection step (`Where(...)`) and the column-selection step (`Query(...)`) compose, instead of the caller having to assemble both into separate inputs to `GetResults(page, shapeids)`.

### 5. Filter shapes by a cell value

**Source:** Synthetic. "Find every shape on the page wider than 2 inches."

**Today:**

```csharp
var page_shapes = page1.Shapes;
var shapeids = Enumerable.Range(0, page_shapes.Count).Select(i => page_shapes[i + 1].ID).ToList();

var query = new VASS.Query.CellQuery();
var col_w = query.Columns.Add(VA.Core.SrcConstants.XFormWidth);
var widths = query.GetResults<double>(page1, shapeids);

var wide_ids = new List<int>();
for (int i = 0; i < shapeids.Count; i++)
{
    if (widths[i][col_w] > 2.0)
        wide_ids.Add(shapeids[i]);
}
```

**With LINQ:**

```csharp
var wide_ids = page1.Shapes
    .Query(s => new { s.ID, s.Width })
    .Results<double>()
    .Where(s => s.Width > 2.0)
    .Select(s => s.ID)
    .ToList();
```

**What this shows:** This is what `Where` is *for*. The today version forces the user to roll their own indexed-zip loop; LINQ takes the predicate as data. **Even with no other LINQ feature, this scenario alone moves a noticeable amount of imperative loop code out of the call site.**

A natural follow-on the spike would have to decide on: should `.Where(s => s.Width > 2.0)` *before* the terminal operator translate to a server-side filter (one COM call returning fewer rows), or is it always a client-side filter on materialized rows? The first matches IQueryable semantics; the second is just LINQ-to-objects. Visio doesn't expose a server-side filter, so likely the second — but that's a real design question.

### 6. Project shapes to a domain record

**Source:** Synthetic. "Dump all shapes' positions and fill colors to JSON."

**Today:**

```csharp
var page_shapes = page1.Shapes;
var shapeids = Enumerable.Range(0, page_shapes.Count).Select(i => page_shapes[i + 1].ID).ToList();

var query = new VASS.Query.CellQuery();
var col_pinx = query.Columns.Add(VA.Core.SrcConstants.XFormPinX);
var col_piny = query.Columns.Add(VA.Core.SrcConstants.XFormPinY);
var col_fg   = query.Columns.Add(VA.Core.SrcConstants.FillForeground);

var formulas = query.GetFormulas(page1, shapeids);
var results  = query.GetResults<double>(page1, shapeids);

var dump = new List<ShapeDump>();
for (int i = 0; i < shapeids.Count; i++)
{
    dump.Add(new ShapeDump {
        Id   = shapeids[i],
        PinX = results[i][col_pinx],
        PinY = results[i][col_piny],
        Fill = formulas[i][col_fg],
    });
}
```

**With LINQ:**

```csharp
var dump = page1.Shapes
    .Query(s => new ShapeDump {
        Id   = s.ID,
        PinX = s.PinX.AsDouble(),
        PinY = s.PinY.AsDouble(),
        Fill = s.FillForeground.AsFormula(),
    })
    .ToList();
```

**What this shows:** Once the user is willing to *name* the projection target (here `ShapeDump`), LINQ goes from "slightly nicer" to "obviously the right shape." The `for` loop and the per-cell index-arithmetic both vanish; the projection literal *is* the spec for what the query reads.

The slight wart: cells that need different result types (formula-as-string for `Fill`, value-as-double for `PinX`/`PinY`) need a way to mark each one. Shown above as `.AsDouble()` / `.AsFormula()` — that's one option; another is to overload by terminal operator (`.Formulas()` vs `.Results<T>()`) at the cost of needing two queries when types are mixed. Spike question.

### 7. Variable-row section (custom properties on many shapes)

**Source:** [`SectionRowHandling`](../../VisioAutomation_2010/VTest/Core/ShapeSheet/ShapeSheetQueryTests.cs#L127). The hardest fit: each shape has 0..N rows of custom properties; the row count varies per shape; each row has multiple cells (Label, Formula, Format, ...).

**Today:**

```csharp
var sec_query = new VASS.Query.SectionQuery();
var sec_cols  = sec_query.Add(IVisio.VisSectionIndices.visSectionProp);
var value_col = sec_cols.Add(VA.Core.SrcConstants.CustomPropValue);
var label_col = sec_cols.Add(VA.Core.SrcConstants.CustomPropLabel);

var shapeidpairs = VA.Core.ShapeIDPairs.FromShapes(shape_a, shape_b, shape_c, shape_d);
var formulas = sec_query.GetFormulas(page1, shapeidpairs);

// formulas[shape_idx][section_idx][row_idx][col_idx]
for (int s = 0; s < formulas.Count; s++)
{
    var shape_sections = formulas[s];
    var prop_section = shape_sections[0];   // section_idx 0 = visSectionProp
    for (int r = 0; r < prop_section.Count; r++)
    {
        string label = prop_section[r][label_col];
        string value = prop_section[r][value_col];
        Console.WriteLine($"shape {s}: {label} = {value}");
    }
}
```

**With LINQ — option A, flatten to row-shaped sequence:**

```csharp
var props = page1.Shapes
    .SelectMany(s => s.CustomProperties.Query(p => new {
        ShapeId = s.ID,
        p.Label,
        p.Formula,
    }))
    .Formulas();

foreach (var p in props)
    Console.WriteLine($"shape {p.ShapeId}: {p.Label} = {p.Formula}");
```

**With LINQ — option B, keep the per-shape grouping:**

```csharp
var by_shape = page1.Shapes
    .Query(s => new {
        ShapeId = s.ID,
        Props   = s.CustomProperties.Select(p => new { p.Label, p.Formula }),
    })
    .Formulas();

foreach (var entry in by_shape)
    foreach (var p in entry.Props)
        Console.WriteLine($"shape {entry.ShapeId}: {p.Label} = {p.Formula}");
```

**What this shows:** This is where the LINQ shape gets tested for real. Option A is the natural LINQ idiom (flatten with `SelectMany`, lose the per-shape grouping) and reads beautifully if the caller actually wants a flat list. Option B keeps the structure but requires the projector to mention a *nested* sequence (`s.CustomProperties.Select(...)`), which is unusual inside a query expression and may not translate cleanly to a single batched COM call.

If neither A nor B can be implemented without either (a) one COM call per shape, or (b) extra round-trips to count rows per section first, then variable-row sections may end up staying on `SectionQuery`. The spike needs to land here decisively.

Note also: the existing `SectionQuery` already pre-counts rows per section per shape via `RowCount[]` (see [`_create_shapesectioncacheitem`](../../VisioAutomation_2010/VisioAutomation/ShapeSheet/Query/SectionQuery.cs#L237)). Any LINQ surface here would need the same machinery; the question is whether wrapping it in `SelectMany` adds enough caller-side ergonomic value to justify the parallel API.

### 8. Compose a query from a runtime list

**Source:** Synthetic, but a real PowerShell pattern. "Caller passed me a list of cell names; query just those."

**Today:**

```csharp
public string[][] QueryByNames(IVisio.Page page, IList<int> shapeids, IEnumerable<string> cell_names)
{
    var dict = ShapeSheetQueryTests.GetSrcDictionary();   // name -> Src

    var query = new VASS.Query.CellQuery();
    var cols = new List<int>();
    foreach (var name in cell_names)
        cols.Add(query.Columns.Add(dict[name]));

    var formulas = query.GetFormulas(page, shapeids);
    var rows = new string[shapeids.Count][];
    for (int s = 0; s < shapeids.Count; s++)
    {
        rows[s] = new string[cols.Count];
        for (int c = 0; c < cols.Count; c++)
            rows[s][c] = formulas[s][cols[c]];
    }
    return rows;
}
```

**With LINQ:**

```csharp
public IEnumerable<Dictionary<string, string>> QueryByNames(
    IVisio.Page page, IList<int> shapeids, IEnumerable<string> cell_names)
{
    var srcs = cell_names.Select(n => SrcDictionary[n]).ToList();

    return page.Shapes
        .Where(s => shapeids.Contains(s.ID))
        .QueryRaw(srcs)
        .Formulas()
        .Select(row => cell_names.Zip(row, (n, v) => (n, v))
                                 .ToDictionary(t => t.n, t => t.v));
}
```

**What this shows:** Composable / runtime-driven queries (the typical PowerShell shape) read about the same in both forms — *if* the LINQ surface exposes a `QueryRaw(IEnumerable<Src>)` escape hatch. Without that escape hatch, callers who don't know cell names at compile time would have to fall back to today's `CellQuery`, which would be a hole in the LINQ story.

This is a nudge to the spike: any LINQ design needs to explicitly answer the dynamic-cell-list case, since one of our two consumer surfaces (PowerShell cmdlets) is *built* on dynamic cell lists.

### 9. The writer side (sketch)

**Source:** [`Formulas_MultipleShapes`](../../VisioAutomation_2010/VTest/Core/ShapeSheet/ShapeSheetWriterTests.cs#L15).

Included as a sketch only — the issue scopes the spike to the *read* side, but it's worth knowing whether LINQ has anything to say about writes.

**Today:**

```csharp
var writer = new VASS.Writers.SidSrcWriter();
writer.SetValue(shape1.ID16, XFormPinX, 0.5);
writer.SetValue(shape1.ID16, XFormPinY, 0.5);
writer.SetValue(shape2.ID16, XFormPinX, 1.5);
writer.SetValue(shape2.ID16, XFormPinY, 1.5);
writer.SetValue(shape3.ID16, XFormPinX, 2.5);
writer.SetValue(shape3.ID16, XFormPinY, 2.5);

writer.Commit(page1, CellValueType.Formula);
```

**With LINQ-flavored writer:**

```csharp
var moves = new[] {
    (shape1, x: 0.5, y: 0.5),
    (shape2, x: 1.5, y: 1.5),
    (shape3, x: 2.5, y: 2.5),
};

moves.Apply(page1, m => new { m.shape.ID16, PinX = m.x, PinY = m.y });
```

**What this shows:** The writer side already has a fairly clean batched-fluent shape (`SetValue` then `Commit`). LINQ-on-writes is a smaller win — closer to "anonymous-record-as-DTO" than "query language as DSL." Probably out of scope for the read-side spike, but worth flagging as an obvious next-question.

## Recap: what the after-side actually buys us

Mapping each scenario to the five frictions in the opening section:

| # | Scenario | Friction(s) addressed | Verdict |
|---|---|---|---|
| 1 | Single typed group, one shape | (1) | Wash. Existing typed records already win here. |
| 2 | Ad-hoc subset, one shape | (1), (3) | Real win. Projection eliminates over-fetch *and* index ceremony. |
| 3 | Mix categories | (1), (2) | Real win. Erases the typed-vs-raw choice. |
| 4 | Many shapes | (1) | Real win. Flat sequence of named records replaces 2D index table. |
| 5 | Filter by cell value | (4) | Real win. Predicate-as-data replaces zip-loop. |
| 6 | Project to domain record | (1), (3), (4) | Big win. The projection literal is the spec. |
| 7 | Variable-row section | (5) | **Open.** Option A reads great if a flat list is what you want; option B is awkward. May not translate to a single batched call. |
| 8 | Runtime cell list | (composability) | Wash, *if* an escape hatch exists. Hole in the design if not. |
| 9 | Writer side | (out of scope) | Smaller win. Defer. |

**The spike's core question**, given the table above, is whether scenarios 2 / 3 / 4 / 5 / 6 — the unambiguous wins — are large enough to justify a parallel API surface alongside today's `CellQuery` / typed `*Cells.cs`, *and* whether scenario 7 can be handled cleanly enough that the LINQ surface is the recommended path for variable-row sections too. If 7 has to fall back to `SectionQuery`, the LINQ story has a visible seam, and that's worth noting in the recommendation.

## Questions to take into the spike

1. **Lazy or eager?** Does the chain `page.Shapes.Where(...).Query(...)` accumulate columns and run one COM call at the terminal operator, or does each operator run its own call? Eager is simpler; lazy is what makes scenarios 4 / 5 / 6 actually batched.
2. **Where does `Where` apply?** Visio has no server-side filter; predicates over cell values must be evaluated client-side after fetching. So `Where(s => s.Width > 2.0)` inevitably fetches `Width` for every shape on the page. Document this clearly or it becomes a perf footgun.
3. **Mixing formula and result projections.** `new { Fill = s.FillForeground, X = s.PinX }` wants `Fill` as `string` (formula) and `X` as `double` (result). Two terminal operators? Per-property markers? Inferred-by-projection-target-type?
4. **Variable-row sections.** Settle on one of (a) `SelectMany` with flat row records, (b) nested `Select` with grouped records, or (c) keep `SectionQuery` and don't try.
5. **Dynamic / runtime cell lists.** Confirm an escape hatch (`QueryRaw(IEnumerable<Src>)` or similar) exists in any final design — PowerShell cmdlet wrappers will need it.
6. **Naming.** `Query` vs `AsShapeSheet` vs `Cells` for the entry point. Cosmetic, but affects discoverability and IntelliSense.
7. **Where does the typed accessor live?** `s.PinX` requires `s` to be *something with* a `PinX` property. Generated record per `*Cells.cs` group? One mega-record? Source-generator? Hand-written facade?

The spike's writeup at [`linq-shapesheet-spike.md`](linq-shapesheet-spike.md) should answer these in the same order.
