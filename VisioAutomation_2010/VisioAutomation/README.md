# VisioAutomation (core library)

The foundation of the solution. Wraps the Microsoft Visio COM API (via `Microsoft.Office.Interop.Visio` v14, the Visio 2010 Primary Interop Assembly) and adds ergonomic .NET types — extension methods on Visio types, batched ShapeSheet I/O, strongly-typed cell groups, and helper facades.

This project is the bottom layer. `VisioAutomation.Models`, `VisioScripting`, and `VisioPowerShell` all depend on it. For the layered architecture see [`docs/ARCHITECTURE.md`](../../docs/ARCHITECTURE.md).

Built into the [`VisioAutomation2010`](https://www.nuget.org/packages/VisioAutomation2010/) NuGet package alongside `VisioAutomation.Models` and `VisioScripting`. Release notes: [`NuGet/CHANGELOG.md`](../../NuGet/CHANGELOG.md).

## Folder layout

- `Application/` — application-level helpers: version detection, UI/window management, undo and alert-response scopes (`ApplicationHelper`, `UndoScope`, `AlertResponseScope`).
- `Core/` — fundamental value types: `Src` and `SidSrc` (ShapeSheet cell addressing), `CellValue`, `Point`, `Size`, `Rectangle`, `ShapeIDPair`.
- `ShapeSheet/` — query and write infrastructure for the Visio ShapeSheet. `Query/CellQuery` builds batched reads; `Writers/SrcWriter` and `SidSrcWriter` build batched writes. Includes metadata caches and stream helpers.
- `Shapes/` — per-shape operations: custom properties, hyperlinks, connection points, controls, locking, geometry. Cell-group types (`ShapeFormatCells`, `ShapeXFormCells`, `LockCells`, …) give strongly-typed views of related cells.
- `Pages/` — page-level helpers and cell groups (`PageFormatCells`, `PageLayoutCells`, `PagePrintCells`, `PageRulerAndGridCells`).
- `Documents/` — document/stencil open and close (`DocumentHelper`).
- `Text/` — character/paragraph cells, tab stops, text formatting.
- `Extensions/` — static extension methods on Visio COM types — the primary fluent API surface (`PageMethods_*`, `ShapeMethods_*`, `MasterMethods_*`, …).
- `Analyzers/` — graph-style analysis of diagrams (`ConnectionAnalyzer` extracts directed edges from connectors).
- `Internal/` — implementation plumbing including `VisioObjectTarget` — a discriminated union over Page/Master/Shape used for unified dispatch.
- `Exceptions/` — `AutomationException`, `VisioOperationException`, `InternalAssertionException`.

## See also

- [`docs/ARCHITECTURE.md`](../../docs/ARCHITECTURE.md) — solution-wide architecture and dependencies
- [`docs/GLOSSARY.md`](../../docs/GLOSSARY.md) — Visio + codebase terminology
- [`docs/BUILDING.md`](../../docs/BUILDING.md) — how to build, test, install
- [`NuGet/CHANGELOG.md`](../../NuGet/CHANGELOG.md) — release notes for the bundled NuGet package
