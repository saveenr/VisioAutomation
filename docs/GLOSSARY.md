# Glossary

Visio-specific and codebase-specific terms used throughout this solution. If you have not worked with Visio's automation model in a while, skim this first — many type names in the code (`Src`, `SidSrc`, `MasterRef`, `ShapeSheet`) only make sense once you know the underlying Visio concepts.

## Visio concepts

**Shape** — Anything drawn on a page: a rectangle, a line, a connector, a grouped sub-diagram, a placed master. Every shape has an `ID` (unique within the page) and a `NameU` (the language-independent name).

**Master** — A reusable shape definition stored in a stencil document. When you "drop" a master onto a page, Visio creates a new shape that references the master. Custom diagram types (org chart node, network device, swimlane) are usually distributed as masters.

**Stencil** — A document that contains masters, typically `.vss`/`.vssx`. Stencils are opened alongside drawing documents to make their masters available.

**Page** — One drawable surface inside a document. A document can have many pages, foreground and background.

**Document** — A Visio file (`.vsd`/`.vsdx`/etc.) containing pages, masters, styles, and document-level settings.

**Connector** — A 1-D shape that links two other shapes, with endpoints that "glue" to connection points. Connectors are the basis of the diagram graph that `ConnectionAnalyzer` extracts.

**ShapeSheet** — A spreadsheet-like grid of cells that backs every Visio object (shape, page, document, even the application). It is the actual source of truth for all formatting, geometry, behavior, and custom data — almost everything you can do via the UI ultimately writes a ShapeSheet cell.

**Cell** — A single addressable value in the ShapeSheet, holding both a *formula* and its evaluated *result*. Examples: `PinX`, `Width`, `LineColor`, `Prop.Owner`. Every cell is identified by a `(Section, Row, Column)` triple.

**Section** — A horizontal grouping of related cells in the ShapeSheet (e.g., the "Custom Properties" section, the "Geometry" section, the "Connection Points" section). Identified by a `VisSectionIndices` constant.

**Row / Column** — Within a section, rows are the records (e.g., one row per custom property) and columns are the fields of that record (label, value, type, …). For non-repeating sections both indices are typically `0`.

**SRC (Section/Row/Cell)** — The `(Section, Row, Cell)` triple that addresses a cell *within a single shape*. In this codebase: [`Src`](../VisioAutomation_2010/VisioAutomation/Core/Src.cs).

**SIDSRC (Shape ID + SRC)** — A `(ShapeID, Section, Row, Cell)` quadruple that addresses a cell on a specific shape *within a page*. In this codebase: [`SidSrc`](../VisioAutomation_2010/VisioAutomation/Core/SidSrc.cs). Used for batched operations across many shapes.

**Formula vs. result** — Every cell holds an editable formula (a string expression like `=Width*0.5`) and an evaluated result (a number, string, or unit-bearing value). The library distinguishes between writing/reading formulas and reading results.

**Custom Property (Shape Data)** — User-defined named fields on a shape, stored in the `Prop.` section of the ShapeSheet. Formerly called "Custom Properties," now "Shape Data" in modern Visio UI.

**User-Defined Cell** — Named cell in the `User.` section, used to hold expressions reused inside other formulas on the same shape.

**PIA (Primary Interop Assembly)** — The .NET assembly that provides managed bindings for a COM library. This solution targets the **Visio 2010 PIA**, `Microsoft.Office.Interop.Visio` v14.

## Codebase concepts

**`VisioObjectTarget`** — An internal discriminated union over Page / Master / Shape. All three host ShapeSheet cells, so helper code uses `VisioObjectTarget` to accept any of them and dispatch correctly. See [`VisioAutomation/Internal/`](../VisioAutomation_2010/VisioAutomation/Internal/).

**`CellQuery` / `SectionQuery`** — Builders that batch multiple cell reads into a single COM call. Add the cells you want, run the query against a shape (or page or master), then index into the results. See [`VisioAutomation/ShapeSheet/Query/`](../VisioAutomation_2010/VisioAutomation/ShapeSheet/Query/).

**`SrcWriter` / `SidSrcWriter`** — Builders that batch multiple cell writes into a single COM call. `SrcWriter` writes to one shape (cells addressed by `Src`); `SidSrcWriter` writes across many shapes on a page (cells addressed by `SidSrc`).

**Cell-group type** — A strongly-typed wrapper around a logically-related set of cells. E.g., `ShapeXFormCells` exposes `PinX`, `PinY`, `Width`, `Height`, `Angle`, `LocPinX`, `LocPinY`. Reading or writing a cell-group type emits a single batched query/write under the hood.

**DOM (in this codebase)** — Not the browser DOM. The declarative diagram tree defined in [`VisioAutomation.Models/Dom/`](../VisioAutomation_2010/VisioAutomation.Models/Dom/): `Document` → `Page` → `Shape` etc. Build it in memory, call `Render(visioApp)` to materialize.

**`MasterRef`** — A reference to a stencil master from the DOM. Resolved against open stencils at render time.

**`Client`** — The entry-point object in `VisioScripting`. Holds a Visio application, a `ClientContext` for output, and command groups as properties.

**`ClientContext`** — Output sink for the scripting layer (`WriteDebug`/`WriteUser`/`WriteError`/`WriteVerbose`/`WriteWarning`). `DefaultClientContext` writes to the console; `VisioPsClientContext` forwards to the PowerShell pipeline.

**`CommandTarget` / `CommandTargetFlags`** — Precondition wrapper. A scripting command declares the state it needs (`RequireApplication`, `RequireDocument`, `RequirePage`); `CommandTarget` resolves and validates that state before the command runs.

**`Target*` types (`TargetPage`, `TargetShapes`, `TargetSelection`, …)** — Deferred-resolution wrappers. `TargetPage.Auto` means *"figure out the active page when the command runs."* This keeps scripts terse by letting callers omit explicit COM objects.
