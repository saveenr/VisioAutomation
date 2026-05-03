# VisioAutomation.Models

A higher-level **declarative document model** for VisioAutomation. Build a tree of plain objects describing the Visio document you want, then call `Render(visioApp)` to materialize it as a real Visio document. Decouples diagram authoring from COM bookkeeping.

Depends on `VisioAutomation` (core). Used by `VisioScripting` and `VisioPowerShell`. For the layered architecture see [`docs/ARCHITECTURE.md`](../../docs/ARCHITECTURE.md).

Built into the [`VisioAutomation2010`](https://www.nuget.org/packages/VisioAutomation2010/) NuGet package alongside `VisioAutomation` and `VisioScripting`. Release notes: [`NuGet/CHANGELOG.md`](../../NuGet/CHANGELOG.md).

> Despite the project name, this isn't a pure POCO assembly — `Render()` methods perform live Visio COM operations.

## Folder layout

- `Dom/` — the core declarative model: `Document` → `PageList` → `Page` → `ShapeList` → `Shape` / `Connector` / `Line` / `Rectangle` / `Oval` / `PolyLine` / `BezierCurve`. Each node has a `Render()` that emits Visio operations. `RenderContext` caches Visio shapes by ID during rendering.
- `Layouts/` — pre-built layouts (Tree, DirectedGraph, Grid, Container, Box) and renderers (`MsaglRenderer`, `VisioLayoutRenderer`) that map graph layouts to a Visio page.
- `LayoutStyles/` — layout strategy types: hierarchy, flowchart, circular, compact tree, radial.
- `Documents/OrgCharts/` — `OrgChartDocument`, a turn-key org-chart generator.
- `Documents/Forms/` — `FormDocument` / `FormPage` / `InteractiveRenderer` for form-style pages.
- `Geometry/` — geometry primitives (`ArcSegment`, `BezierSegment`, `LineSegment`, `BezierCurve`, `BoundingBoxBuilder`).
- `Color/` — `ColorRgb`, `ColorHsl`.
- `Text/` — rich text model (`Element`, `Literal`, `Field`, character/paragraph formatting, regions).
- `Data/` — adapters from `DataTable` / XML to diagram structure.
- `Utilities/` — `MasterCache` and similar helpers.

## External dependencies (beyond `VisioAutomation` core)

- [Microsoft.Msagl](https://www.nuget.org/packages/Microsoft.Automatic.Graph.Layout) — Microsoft Automatic Graph Layout, used by the layout engines.
- [GenTreeOps](https://www.nuget.org/packages/GenTreeOps) — tree-manipulation utilities.

## See also

- [`docs/ARCHITECTURE.md`](../../docs/ARCHITECTURE.md) — solution-wide architecture and dependencies
- [`docs/GLOSSARY.md`](../../docs/GLOSSARY.md) — Visio + codebase terminology
- [`docs/BUILDING.md`](../../docs/BUILDING.md) — how to build, test, install
- [`NuGet/CHANGELOG.md`](../../NuGet/CHANGELOG.md) — release notes for the bundled NuGet package
