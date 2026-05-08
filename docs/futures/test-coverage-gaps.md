# Test coverage gaps (audit)

Output of [#154](https://github.com/saveenr/VisioAutomation/issues/154). A pass over the public API surface against the existing test inventory, producing a prioritized gap list. **Audit only**: no test-writing happens in this doc; closing the gaps is follow-up work. For backlog context see [`tests.md`](tests.md).

## Method

Public surface enumerated by grepping `^\s*public\s+(static\s+)?(sealed\s+)?(abstract\s+)?(partial\s+)?(class|struct|enum|interface|record)\s+\w+` against each source project, excluding `Internal/`. Test methods enumerated by grepping `\[(MUT\.)?TestMethod\]` against each test project. Cross-reference is by symbol name and namespace mapping. The audit identifies *which types and helpers have no direct unit/integration test*, not whether existing tests are deep enough — e.g. `ShapeHelper` has 3 tests all targeting `GetNestedShapes`, so the rest of `ShapeHelper`'s public surface is technically untested even though the type itself appears in the test inventory.

## Inventory snapshot (2026-05-07)

| Project | Public types (approx.) | Test methods | Test files |
|---|---|---|---|
| `VisioAutomation` | ~105 | 97 | 27 |
| `VisioAutomation.Models` | ~104 | 60 | 12 |
| `VisioScripting` | ~64 | 34 | 19 |
| `VisioPowerShell/Commands` | ~70 cmdlets | 20 | 2 |
| **Total** | **~340** | **211** | **60** |

The ~211 enumerated method count is slightly under the 215-test runtime count from the last green run; the discrepancy is attributable to test methods whose `[TestMethod]` attribute and `public void` signature span multiple lines (the regex matches single-line forms only). Close enough for an audit.

## Cross-cutting observations

1. **PowerShell cmdlet coverage is the single biggest gap.** 70 cmdlets ship in the module; `VTest.PowerShell\VisioPS_Basic_Tests` covers ~14 of them with functional tests, plus 2 manifest-export tests. The remaining ~56 cmdlets have no dedicated test. Their underlying business logic generally lives in `VisioScripting` command-set classes that *are* tested at the .NET level, but cmdlet-level tests catch a distinct regression class — parameter binding (positional, mandatory, switches), `ParameterSetName` resolution, pipeline binding, output piping. **The four bug fixes shipped in PowerShell module 4.6.1 (Lock/Unlock switch binding, Export inverted file-existence check, `New-VisioShape` polyline minimum-point validation) are exactly the regression class that cmdlet-level tests would have caught.** `VisioPS_Basic_Tests` already has the harness (`VisioPS_Session.InvokeScript<T>`); the gap is fixture-style coverage breadth, not infrastructure.

2. **`VisioScripting` command-set tests are broad but shallow.** 19 of the ~30 `*Commands` classes have a dedicated test file in `VTest.Scripting`, mostly with 1-3 happy-path tests each. Edge cases (multi-shape selection, empty selection, no-active-document, error paths, wildcards) are largely untested. Eight command-set classes have no dedicated test file at all (LayerCommands, LockCommands, ModelCommands, OutputCommands, UndoCommands, UserDefinedCellCommands, ViewCommands; ContainerCommands has only one DropContainer scenario).

3. **`VisioAutomation.Extensions/*Methods.cs` static method bags are mostly untested directly.** ~22 method-bag classes (`*Methods.cs` under `Extensions/`); about 6 have any direct test (`PageMethods_Draw`, `PageMethods_Drop`, `PageMethods_General`, `SelectionMethods`, `DocumentMethods.ForceClose`, `ConnectsMethods` via `Connects_EnumerableExtensionMethod`), and those that do typically cover one method out of many. They are exercised transitively via the helpers and command sets that delegate to them, but no direct unit-level coverage exists for the rest.

4. **`Cells` record classes (`*Cells.cs` deriving from `CellRecord`) are exercised indirectly but rarely directly.** ~17 cell-record types across `VisioAutomation/Shapes/`, `VisioAutomation/Pages/`, `VisioAutomation/Text/`. None have round-trip-through-formula-and-back tests at the type level. They are validated transitively whenever a helper reads/writes them, which is most of the time, so the gap is real but lower priority.

5. **The `Documents/Forms/` package in `VisioAutomation.Models` has no tests.** `FormDocument`, `FormPage`, `InteractiveRenderer`, `TextBlock`, `PageMargin` — no entries in `VTest.Models`. The Tier 3 doc audit added a Forms gitbook page on 2026-05-07 (`models/forms`) with a working snippet adapted from sample code, but no test pinned that snippet. This is the only sub-package of `VisioAutomation.Models` with zero direct coverage.

## Well-covered areas (no gap)

For orientation. These are areas where a reasonable reader looking for "what's tested" would find enough:

- `VisioAutomation.Analyzers.ConnectionAnalyzer` and `Path` — 11 tests covering edge-direction-source variants, no-arrow handling, transitive closure.
- `VisioAutomation.Application.UndoScope` — 6 tests covering simple, nested, abort, and abort-nested scenarios.
- `VisioAutomation.ShapeSheet.Writers` and `Query` — 13 tests covering single/multi-shape, formulas, results-int/string/double, consistency checking, section-row handling, out-of-order verification.
- `VisioAutomation.Shapes.UserDefinedCellHelper` and `CustomPropertyHelper` — 21 tests across the two helpers covering get/set, multi-shape, name validation, value quoting, encoded-vs-raw characterization, typed-setters round-trip.
- `VisioAutomation.Models.DOM` core types (Document, Page, Shape, Connector, ShapeList, ...) — 17 tests in `DOM_Tests.cs`.
- `VisioAutomation.Models.Layouts.DirectedGraph` — 22 tests in `DrawModel_DirectedGraph.cs` covering renderer + MSAGL options + node sizing (incl. the [#82](https://github.com/saveenr/VisioAutomation/issues/82) regression).
- `VisioAutomation.Models.Documents.OrgCharts` — 5 tests covering layout direction, styling, hierarchy.

## Gap list — High priority

Public API, no direct test, and behaviorally meaningful. Closing these gaps catches real regressions; closing the others is hygiene.

### `VisioPowerShell/Commands` (the big one)

The 56 cmdlets without a `VTest.PowerShell\VisioPS_Basic_Tests` test, grouped by surface area. Suggested coverage shape per cmdlet: 1-2 tests asserting parameter binding (positional / pipeline / switch) plus a happy-path execution returning a non-null result. Heavier behavioral coverage already lives in `VTest.Scripting`.

| Surface | Cmdlets without dedicated tests |
|---|---|
| VisioApplication | `Close-VisioApplication`, `New-VisioApplication`, `Out-Visio`, `Redo-Visio`, `Test-VisioApplication`, `Undo-Visio` |
| VisioControl | `New-VisioControl`, `Remove-VisioControl` |
| VisioCustomProperty | `Remove-VisioCustomProperty`, `Set-VisioCustomProperty` |
| VisioDocument | `Close-VisioDocument`, `New-VisioDocument`, `Open-VisioDocument`, `Save-VisioDocument`, `Select-VisioDocument`, `Test-VisioDocument` |
| VisioHyperlink | `New-VisioHyperlink`, `Remove-VisioHyperlink` |
| VisioModel | `Import-VisioModel` |
| VisioPage | `Copy-VisioPage`, `Export-VisioPage`, `Format-VisioPage`, `Measure-VisioPage`, `New-VisioPage`, `Remove-VisioPage`, `Select-VisioPage` |
| VisioPageCells | `New-VisioPageCells` |
| VisioPoint | `New-VisioPoint` |
| VisioRectangle | `New-VisioRectangle` |
| VisioShape | `Connect-VisioShape`, `Copy-VisioShape`, `Export-VisioShape`, `Format-VisioShape`, `Join-VisioShape`, `Lock-VisioShape`, `Measure-VisioShape`, `New-VisioShape`, `Remove-VisioShape`, `Select-VisioShape`, `Split-VisioShape`, `Test-VisioShape`, `Unlock-VisioShape` |
| VisioShapeCells | `New-VisioShapeCells` |
| VisioText | `Set-VisioText` |
| VisioUserDefinedCell | `Remove-VisioUserDefinedCell` |
| VisioWindow | `Format-VisioWindow` |

The four cmdlets fixed in 4.6.1 — `Lock-VisioShape`, `Unlock-VisioShape`, `Export-VisioShape`, `New-VisioShape` — and `Connect-VisioShape` (pipeline parameter set, [#163](https://github.com/saveenr/VisioAutomation/issues/163), parked behind [#164](https://github.com/saveenr/VisioAutomation/issues/164)) are the highest-priority targets within this list. Their bug histories are evidence the binding surface needs guarding.

### `VisioScripting/Commands` (no dedicated test file)

Public command-set classes consumed by cmdlets, no dedicated `Scripting_*Tests.cs` file:

- `LayerCommands` — layer add/remove/membership; `VisioPowerShell/Commands/VisioShape/JoinVisioShape.cs` etc. ride on this.
- `LockCommands` — direct backing for `Lock-VisioShape` / `Unlock-VisioShape`. Adding tests here also helps secure the 4.6.1 fix surface.
- `ModelCommands` — backing for `Import-VisioModel`. The DOM/DirectedGraph/OrgChart drawing flows in `VTest.Models` exercise this transitively but no direct test pins the command-set surface.
- `OutputCommands` — backing for `Out-Visio`. Probably small surface but currently unverified.
- `UndoCommands` — `client.Undo.UndoLastAction()` is touched once in `Scripting_ApplicationTests`, but undo/redo at the command-set level (incl. `BeginUndoScope` / `EndUndoScope` shapes) is not directly tested.
- `UserDefinedCellCommands` — `VisioAutomation.Shapes.UserDefinedCellHelper` is well tested; the scripting-layer wrapper that aggregates over `TargetShapes` is not.
- `ViewCommands` — zoom / view-mode scripting; surface is small but unguarded.

### `VisioAutomation.Models.Documents.Forms` (whole sub-package)

`FormDocument`, `FormPage`, `InteractiveRenderer`, `TextBlock`, `PageMargin`. No tests in `VTest.Models`. Suggested: one round-trip test that builds a small `FormDocument`, calls `InteractiveRenderer`, and asserts the rendered shape count and page bounds.

### `VisioAutomation.Application.AlertResponseScope` and `UserInterfaceHelper`

`AlertResponseScope` suppresses Visio modal dialogs during automation; a test would set up a scenario that would normally prompt and assert non-prompting + correct fallback response code. `UserInterfaceHelper` covers UI-mode toggling and probably doesn't need much, but an instantiation/round-trip test on at least one method would close the gap.

### `VisioScripting/Helpers/WildcardHelper`

Pattern-matching used across cmdlets with `*` / `?` patterns — a regression here silently changes which shapes/pages get matched. Currently no direct test. This is a pure-CPU helper, so the test cost is near zero.

## Gap list — Medium priority

Partial coverage: the type appears in the test inventory but only one or two methods are tested, leaving most of the surface unguarded.

### `VisioAutomation` (helpers with 1-2 tests where the surface is much larger)

- `VisioAutomation.Application.ApplicationHelper` — 1 test (`TestStencilLocation`). Surface includes app-startup, document-collection ops, and stencil-cache logic.
- `VisioAutomation.Shapes.ShapeHelper` — 3 tests, all `GetNestedShapes` variants. Other public methods (group-list traversal, master inheritance) are untested.
- `VisioAutomation.Shapes.HyperlinkHelper` — 1 test (`Hyperlinks_AddRemove`). Edit / sub-address / keep-hyperlinks scenarios untested.
- `VisioAutomation.Shapes.ConnectorHelper` — 1 test (`Connect1`). Multi-connect, cross-page, and dynamic-connector scenarios untested.
- `VisioAutomation.Shapes.ControlHelper` — 1 test (`Controls_AddRemove`). Position-set and visibility-cell scenarios untested.
- `VisioAutomation.Shapes.GeometryHelper` — 2 tests (`Geometry_AddGeometrySection`, `Geometry_DeleteGeometry`). Per-row mutation and multi-section handling untested.
- `VisioAutomation.Text.TextHelper` — 1 test (`Text_Case1`). Field-encode-decode, text-runs, and tab-stop construction surface untested.
- `VisioAutomation.Text.TextFormat` — 1 test (`Text_TabStops_Set`). Character-cells round-trip and paragraph-cells round-trip untested.
- `VisioAutomation.Pages.PageHelper` — 5 tests covering Query/Orientation/Duplicate/SwitchPages/ResizeBorder. Background-page assignment and reorder untested.
- `VisioAutomation.Documents.DocumentHelper` — appears untested directly (the related `DocumentMethods.ForceClose` is the one tested touch point). Stencil-load and template-load paths untested.

### `VisioAutomation.Extensions/*Methods.cs`

The static method-bag classes are mostly untested directly:

- **No direct test:** `ColorsMethods`, `FontsMethods`, `LayersMethods`, `MasterMethods_Drop`, `MasterMethods_Draw`, `MasterMethods_General`, `MasterMethods_ShapeSheet`, `PageMethods_ShapeSheet`, `SectionMethods`, `ShapeMethods_Draw`, `ShapeMethods_Drop`, `ShapeMethods_General`, `ShapeMethods_ShapeSheet`, `StylesMethods`, `WindowMethods`.
- **One method touched:** `ConnectsMethods` (via `Connects_EnumerableExtensionMethod`), `DocumentMethods` (via `Document_ForceClose`), `PageMethods_Drop` (via `Page_Drop_ManyU`), `PageMethods_General` (via `Page_CreatePage`), `SelectionMethods` (via `Selection_GetShapeIDs` / `Selection_ToEnumerable`).

These are exercised transitively, so the priority here is medium not high. Still, the *number* of untouched static methods across these classes is large — closing this gap is a useful follow-up project on its own.

### `VisioAutomation.Models.Layouts.Grid` and `Layouts.Tree`

- `GridLayout` — 1 test (`DrawModel_Grid.Scripting_Draw_DataTable` analog). Column-direction, row-direction, mixed-orientation untested.
- `TreeLayout` — 2 tests (`Tree_Tests`: SingleNode, MultiNode). Layout-direction variants, connector-type variants, orientation untested.

For comparison: `Layouts.DirectedGraph` (similar complexity) has 22 tests. Grid and Tree are under-covered relative to their public surface.

### `VisioAutomation.Models.LayoutStyles`

12 styling-config classes (`CircularLayoutStyle`, `CompactTreeLayout`, `FlowchartLayoutStyle`, `HierarchyLayoutStyle`, `RadialLayoutStyle`, plus enum companions). Tested only insofar as a `DrawModel_*` test happens to use one. No dedicated style-construction or style-effect tests.

### `VisioAutomation.Models.Text` (DOM-side text formatting)

12 text-tree types (`Field`, `CustomField`, `FieldBase`, `Element`, `Literal`, `CharacterFormatting`, `ParagraphFormatting`, `NodeList<T>`, `Node`, `NodeType`, `CharStyle`, `FieldConstants`). 1 dedicated test (`Dom_Text_Tests`). The model side of the rich-text DOM is significantly under-covered.

### `VisioScripting` core types (target/wildcard resolution)

`Client`, `ClientContext`, `CommandTarget`, `TargetShapes`, `TargetPages`, `TargetDocuments` — well-exercised transitively (every `VTest.Scripting` test routes through them). No direct unit tests for resolution logic specifically: `TargetShapes.None` semantics, wildcard expansion, "no active document" fallthroughs, multi-window targeting. A regression in resolution would surface as inconsistent failures across the broader test surface, which is hard to triage.

## Gap list — Low priority

Internal helpers, data-shape-only types, enums, and value types whose surface is small enough or whose indirect coverage is good enough that direct tests give marginal value.

### Data-shape `Cells` records (round-trip tests would be hygiene only)

- `VisioAutomation.Pages.{PageFormatCells, PageRulerAndGridCells, PagePrintCells, PageLayoutCells}`
- `VisioAutomation.Shapes.{ShapeFormatCells, ShapeLayoutCells, ShapeXFormCells, ControlCells, ConnectionPointCells, HyperlinkCells, LockCells, CustomPropertyCells, UserDefinedCellCells}`
- `VisioAutomation.Text.{CharacterCells, ParagraphCells, TextBlockCells, TextXFormCells}`
- `VisioAutomation.Models.DOM.{ShapeCells, PageCells}`

These are predominantly field-bag classes that map ShapeSheet cells to typed properties. Indirectly exercised whenever a helper reads or writes them.

### Value types and enums

Tested only when their parent type is tested:

- `VisioAutomation.Core.{CellValue, CellValueType, SidSrc, Src, SrcConstants, ShapeIDPairs}` — `SidSrc` and `Src` size are tested. `CellValue` round-trip would be a small useful addition; `ShapeIDPairs` is rarely materialized outside Connect operations.
- `VisioAutomation.Analyzers.{EdgeNoArrowsHandling, EdgeDirectionSource, DirectedEdge, BitArray2D}` — covered transitively via `ConnectionAnalysisTests` and `Construct2DBitArray`.
- `VisioAutomation.Application.{AlertResponseCode}` — enum companion to `AlertResponseScope`.
- `VisioAutomation.Models.Color.{ColorRgb, ColorHsl}` — color value structs. Construction-and-equality tests would be quick.
- `VisioAutomation.Models.Geometry.{LineSegment, BezierSegment}` — value types. The corresponding `BoundingBoxBuilder` and `BezierCurve` *are* tested.
- `VisioScripting/Models/{AlignmentHorizontal, AlignmentVertical, Axis, ConnectionPointType, PageOrientation, PageRelativePosition, SelectionOperation, ShapeSelectionOperation, ShapeSendDirection, ZoomToObject}` — all enums.

### Internal-feeling helpers and utility types

- `VisioAutomation.Application.Logging.{LogRecord, LogSession, LoggingHelper}` — `XmlErrorLog` is tested; the surrounding logging plumbing isn't. `XmlErrorLog` is the user-facing entry point, which is what matters.
- `VisioAutomation.Core.BasicList<T>` — base class; tested transitively by every consumer.
- `VisioAutomation.Exceptions.{AutomationException, VisioOperationException, InternalAssertionException}` — exception types with no custom logic.
- `VisioAutomation.Models.Data.XmlModel` — used internally by `DataTableModel`; tested transitively.
- `VisioScripting/Helpers/{InteropHelper, ReflectionHelper, SelectionHelper}` — small reflection/coercion helpers; tested transitively. `ReflectionHelper.NamingOptions` is consumed by `DeveloperCommands` which has 1 test.
- `VisioScripting/Models/{EnumType, EnumValue, ShapeSheetReader, ShapeSheetWriter, PageDimensions, ShapeDimensions}` — `ShapeSheetReader` / `ShapeSheetWriter` are non-trivial; tested transitively via `Scripting_ShapeSheetTests`.
- `VisioScripting/Loaders/{OrgChartDocumentLoader, DirectedGraphDocumentLoader}` — covered transitively by `VTest.Models.DrawModel_OrgChartTests` (5) and `DrawModel_DirectedGraph` (22).
- `VisioAutomation.ShapeSheet.Streams.{StreamArray, StreamType}` — covered transitively by writer tests.
- `VisioAutomation.ShapeSheet.Data.{DataColumns, DataColumn, DataRows<T>, DataRow<T>, DataRowGroups<T>, DataRowGroup<T>}` — covered transitively by `SectionQuery` tests.
- `VisioAutomation.ShapeSheet.CellRecords.{CellRecord, CellRecords<T>, CellRecordsGroup<T>, CellRecordBuilderCellQuery<T>, CellRecordBuilderSectionQuery<T>}` — base types for the cell-record hierarchy; tested transitively.
- `VisioAutomation.Models.Layouts.InternalTree.{AlignmentVertical, ParentChildConnection<U>}` — internal helpers for `TreeLayout`.
- `VisioAutomation.Models.DOM.{NodeList<T>, Node, BaseShape, MasterRef, RenderPerformanceSettings, Hyperlink}` — DOM scaffolding; tested transitively via `DOM_Tests`.
- `VisioAutomation.Models.Documents.OrgCharts.{Node, NodeList}` — covered transitively by `DrawModel_OrgChartTests`.

## Suggested first slice for follow-up implementation

If this list is pulled into actual test-writing work, the highest leverage cuts are:

1. **PowerShell cmdlet binding tests for `Lock-`, `Unlock-`, `Export-VisioShape`, `New-VisioShape`, and `Connect-VisioShape`.** Five cmdlets, ~10 tests, directly guards the regression class that shipped in 4.6.1 and unblocks closing [#163](https://github.com/saveenr/VisioAutomation/issues/163) once [#164](https://github.com/saveenr/VisioAutomation/issues/164) is decided.
2. **A `Scripting_LockTests.cs` and `Scripting_LayerTests.cs` in `VTest.Scripting`.** Two new files at the same shape as the existing 19, ~5 tests total, covers two of the eight fully-untested command sets and reinforces (1).
3. **A `VTest.Models\FormsTests.cs`.** Closes the only sub-package of `VisioAutomation.Models` with zero coverage and pins the gitbook `models/forms` snippet.

The remainder of the list is hygiene: each item is small but the aggregate is open-ended. Any follow-up implementation issue should pick a cap (e.g. "add 30-50 tests targeting the high-priority list") rather than chase the whole audit.
