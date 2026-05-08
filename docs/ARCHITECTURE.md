# VisioAutomation — Architecture

This document describes the structure of the VisioAutomation solution: what each project is for, how the projects depend on one another, and the central concepts a developer needs to know to navigate the code.

For a glossary of Visio-specific terms (ShapeSheet, SRC, master, etc.) see [GLOSSARY.md](GLOSSARY.md). For build and test instructions see [BUILDING.md](BUILDING.md).

---

## 1. What this solution is

**VisioAutomation** is a set of .NET libraries that wrap the Microsoft Visio COM automation API (`Microsoft.Office.Interop.Visio`, the Visio 2010 PIA) and expose it as ergonomic, strongly-typed .NET APIs. On top of the wrapper sits a higher-level scripting facade and a PowerShell module front-end.

The solution is focused on **out-of-process automation** of a running Visio instance — it is not a Visio add-in framework, and it does not render Visio diagrams without Visio.

The solution file is [`VisioAutomation_2010/VisioAutomation2010.sln`](../VisioAutomation_2010/VisioAutomation2010.sln) and contains 10 projects.

---

## 2. Layered architecture

The solution has four production layers and three layers of tests/samples. Dependencies flow strictly downward.

```
┌─────────────────────────────────────────────────────────────┐
│  VisioPowerShell        (PowerShell module: Visio.psd1)     │
│  Cmdlets like New-VisioShape, Get-VisioPage                 │
└─────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────┐
│  VisioScripting         (Client + verb-noun command groups) │
│  Client.Page.SetActivePage(...), Client.Draw.AddShape(...)  │
└─────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────┐
│  VisioAutomation.Models (Declarative DOM + layout engines)  │
│  Document → PageList → ShapeList; OrgChart, Tree, Grid, ... │
└─────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────┐
│  VisioAutomation        (Core: ShapeSheet, helpers, exts)   │
│  CellQuery, SrcWriter, ShapeHelper, PageMethods_*           │
└─────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────┐
│  Microsoft.Office.Interop.Visio v14 (Visio 2010 PIA, COM)   │
└─────────────────────────────────────────────────────────────┘
```

Auxiliary external dependencies (used by the Models layer): **Microsoft.Msagl** (graph layout) and **GenTreeOps** (tree utilities).

---

## 3. Projects

### 3.1 Production projects

#### `VisioAutomation` — the core wrapper
Path: [`VisioAutomation_2010/VisioAutomation/`](../VisioAutomation_2010/VisioAutomation/) · TFM: .NET Framework 4.5.2

The foundation of the solution. Wraps the Visio COM API and adds:

| Folder | Role |
|---|---|
| `Application/` | Application-level helpers — version detection, UI/window management, undo and alert-response scopes (`ApplicationHelper`, `UndoScope`, `AlertResponseScope`). |
| `Core/` | Fundamental value types: `Src` and `SidSrc` (ShapeSheet cell addressing), `CellValue`, `Point`, `Size`, `Rectangle`, `ShapeIDPair`. |
| `ShapeSheet/` | Query and write infrastructure for the Visio ShapeSheet. `Query/CellQuery` builds batch reads; `Writers/SrcWriter` and `SidSrcWriter` build batch writes. Includes metadata caches (`ShapeMetadataCache`, `SectionMetadataCache`) and stream helpers. |
| `Shapes/` | Per-shape operations: custom properties, hyperlinks, connection points, controls, locking, geometry. Cell-group types (`ShapeFormatCells`, `ShapeXFormCells`, `LockCells`, …) give strongly-typed views of related ShapeSheet cells. |
| `Pages/` | Page-level helpers and cell groups (`PageFormatCells`, `PageLayoutCells`, `PagePrintCells`, `PageRulerAndGridCells`). |
| `Documents/` | Document/stencil open and close (`DocumentHelper`). |
| `Text/` | Character/paragraph cells, tab stops, text formatting. |
| `Extensions/` | Static extension methods on Visio COM types — the primary fluent API surface (`PageMethods_*`, `ShapeMethods_*`, `MasterMethods_*`, `ApplicationMethods`, etc.). |
| `Analyzers/` | Graph-style analysis of diagrams (`ConnectionAnalyzer` extracts directed edges from connectors). |
| `Internal/` | Implementation plumbing including `VisioObjectTarget` — a discriminated union over Page/Master/Shape used to dispatch operations through unified code paths. |
| `Exceptions/` | `AutomationException`, `VisioOperationException`, `InternalAssertionException`. |

**External deps:** `Microsoft.Office.Interop.Visio` v14 (Visio 2010 PIA, via the `Visio2010.PrimaryInteropAssembly` NuGet package). No other NuGet packages.

**Dominant patterns:** extension methods on Visio COM types; fluent batch query/write builders for the ShapeSheet; cell-group types that give strongly-typed views of related cells; metadata caching to avoid repeated COM round trips.

---

#### `VisioAutomation.Models` — declarative DOM and layouts
Path: [`VisioAutomation_2010/VisioAutomation.Models/`](../VisioAutomation_2010/VisioAutomation.Models/) · TFM: .NET Framework 4.5.2

A higher-level **declarative document model**. Build a tree of plain objects describing the diagram you want, then call `Render(visioApp)` to materialize it as a real Visio document. This decouples diagram authoring from COM bookkeeping.

| Folder | Role |
|---|---|
| `Dom/` | The core declarative model: `Document` → `PageList` → `Page` → `ShapeList` → `Shape`/`Connector`/`Line`/`Rectangle`/`Oval`/`PolyLine`/`BezierCurve`. Each node has a `Render()` that emits Visio COM operations. `RenderContext` caches Visio shapes by ID during rendering. |
| `Layouts/` | Pre-built layouts (Tree, DirectedGraph, Grid, Container, Box) and renderers (`MsaglRenderer`, `VisioLayoutRenderer`) that map graph layouts to a Visio page. |
| `LayoutStyles/` | Layout strategy types: hierarchy, flowchart, circular, compact tree, radial. |
| `Documents/OrgCharts/` | `OrgChartDocument` — turn-key org-chart generator. |
| `Documents/Forms/` | `FormDocument`/`FormPage`/`InteractiveRenderer` for form-style pages. |
| `Geometry/` | Geometry primitives — `ArcSegment`, `BezierSegment`, `LineSegment`, `BezierCurve`, `BoundingBoxBuilder`. |
| `Color/` | `ColorRgb`, `ColorHsl`. |
| `Text/` | Rich text model — `Element`, `Literal`, `Field`, character/paragraph formatting, regions. |
| `Data/` | Adapters from `DataTable` / XML to diagram structure. |
| `Utilities/` | `MasterCache` and other helpers. |

**External deps:** Visio 2010 PIA, `Microsoft.Msagl` (graph layout), `GenTreeOps` (tree ops). **Project ref:** `VisioAutomation`.

**Note:** despite the name, Models is not a pure POCO assembly — its `Render()` methods perform live Visio COM operations.

---

#### `VisioScripting` — high-level facade
Path: [`VisioAutomation_2010/VisioScripting/`](../VisioAutomation_2010/VisioScripting/) · TFM: .NET Framework 4.5.2

A scripting-friendly facade over the core library. Organizes operations into ~25 verb-noun **command groups** hung off a single [`Client`](../VisioAutomation_2010/VisioScripting/Client.cs) instance.

**Key types:**

- [`Client`](../VisioAutomation_2010/VisioScripting/Client.cs) — entry point, constructed with an `IVisio.Application`. Exposes command groups as properties (`Application`, `Document`, `Page`, `Selection`, `Draw`, `Text`, `Arrange`, `Connection`, `ShapeSheet`, `Layer`, `Grouping`, `Master`, `CustomProperty`, `Hyperlink`, `Control`, `View`, `Export`, `Undo`, …).
- [`ClientContext`](../VisioAutomation_2010/VisioScripting/ClientContext.cs) — abstract output sink: `WriteDebug`/`WriteUser`/`WriteError`/`WriteVerbose`/`WriteWarning`. Subclass to redirect logging.
- [`DefaultClientContext`](../VisioAutomation_2010/VisioScripting/DefaultClientContext.cs) — writes to `Console`.
- [`CommandTarget`](../VisioAutomation_2010/VisioScripting/CommandTarget.cs) + [`CommandTargetFlags`](../VisioAutomation_2010/VisioScripting/CommandTargetFlags.cs) — preconditions wrapper. A command declares it requires (e.g.) an active page, and `CommandTarget` validates and resolves that state up front.
- `Target*` family (`TargetDocument`, `TargetPage`, `TargetShapes`, `TargetSelection`, `TargetWindow`, `TargetPages`, `TargetDocuments`, `TargetObjects`) — deferred-resolution wrappers. `TargetPage.Auto` means "use the active page when this command runs," so callers rarely have to pass explicit COM objects.

**Project refs:** `VisioAutomation`, `VisioAutomation.Models`.

**Typical use:**
```csharp
var client = new Client(visioApp);
client.Document.NewDocument();
client.Draw.DrawRectangle(TargetPage.Auto, new Rectangle(0, 0, 4, 2));
client.Text.SetText(TargetShapes.Auto, "Hello");
```

---

#### `VisioPowerShell` — PowerShell module
Path: [`VisioAutomation_2010/VisioPowerShell/`](../VisioAutomation_2010/VisioPowerShell/) · TFM: .NET Framework 4.5.2

Binary PowerShell module (`VisioPS.dll`) shipped as the `Visio` module. Module manifest: [`Visio.psd1`](../VisioAutomation_2010/VisioPowerShell/Visio.psd1) (version 4.6.0). Cmdlets follow strict verb-noun naming (`Get-VisioShape`, `New-VisioShape`, `Set-VisioShapeCells`, `Select-VisioShape`, `Close-VisioApplication`, …) and are organized under `Commands/` by noun (`VisioApplication/`, `VisioDocument/`, `VisioPage/`, `VisioShape/`, `VisioShapeCells/`, `VisioPageCells/`, `VisioControl/`, `VisioCustomProperty/`, `VisioHyperlink/`, `VisioUserDefinedCell/`, `VisioText/`, `VisioMaster/`, `VisioWindow/`, …).

**How a cmdlet reaches Visio:**

1. Each cmdlet inherits from `VisioCmdlet`, which holds a session-scoped `VisioScripting.Client`.
2. Cmdlet output and logging are bridged through [`VisioPsClientContext`](../VisioAutomation_2010/VisioPowerShell/VisioPsClientContext.cs), a `ClientContext` subclass that forwards `WriteVerbose`/`WriteDebug`/etc. to PowerShell's pipeline.
3. The cmdlet's `ProcessRecord()` calls into `this.Client.<Group>.<Method>(...)`.

[`Visio.Types.ps1xml`](../VisioAutomation_2010/VisioPowerShell/Visio.Types.ps1xml) customizes how Visio COM objects display in PowerShell (e.g., `Shape` shows `NameU`, `ID`, `Type`).

**Project refs:** `VisioScripting`, `VisioAutomation.Models`, `VisioAutomation`. **External:** `System.Management.Automation` v3, Visio 2010 PIA.

**Loader scripts:**
- [`LoadFromBinDebug.ps1`](../VisioAutomation_2010/VisioPowerShell/LoadFromBinDebug.ps1) — import the freshly built debug DLL for fast iteration.
- [`LoadFromBinDebug.ISE.ps1`](../VisioAutomation_2010/VisioPowerShell/LoadFromBinDebug.ISE.ps1) — same, but launched inside the PowerShell ISE.
- [`LoadFromGallery.ps1`](../VisioAutomation_2010/VisioPowerShell/LoadFromGallery.ps1) — `Save-Module` the published Visio module from PSGallery into a local subfolder and import from there. Useful for release-verification against the published artifact.
- [`InstallForCurrentUser.ps1`](../VisioAutomation_2010/VisioPowerShell/InstallForCurrentUser.ps1) — robocopy the built artifacts into the user's PowerShell modules directory.

---

### 3.2 Test projects

All test projects use **MSTest** (`MSTest.TestFramework` 4.2.2) and **require a live, locally-installed Microsoft Visio** because they exercise real COM calls.

| Project | TFM | Tests |
|---|---|---|
| [`VTest`](../VisioAutomation_2010/VTest/) | .NET 4.7.2 | Core library — ShapeSheet, geometry, connectors, hyperlinks, custom properties, analyzers. The `Framework/` folder owns shared infrastructure: `VTest` base class, `VTestAppRef` (Visio COM lifecycle), `VTestHelper`, `VTestScriptingClient`. |
| [`VTest.Models`](../VisioAutomation_2010/VTest.Models/) | .NET 4.7.2 | DOM + layouts — `OrgChartDrawModelTests`, `DirectedGraphDrawModelTests`, `GridDrawModelTests`, `BoxLayoutTests`, `TreeTests`, `BezierTests`. |
| [`VTest.Scripting`](../VisioAutomation_2010/VTest.Scripting/) | .NET 4.7.2 | Scripting facade — `ApplicationTests`, `DocumentTests`, `ShapeSheetTests`, `ExportTests`, `GroupTests`, `PageTests`. |
| [`VTest.PowerShell`](../VisioAutomation_2010/VTest.PowerShell/) | .NET 4.7.2 | PowerShell cmdlet integration — uses `VTestPowerShellSession` to spin up an in-process PS session and execute cmdlets against live Visio. |

> Note the TFM mismatch: the production libraries target .NET 4.0 but several test projects target .NET 4.7.2. This is a known cleanup item for the 2026 refresh.

---

### 3.3 Sample / demo projects

| Project | What it is |
|---|---|
| [`VSamples`](../VisioAutomation_2010/VSamples/) | WinForms sample runner (.NET 4.0). ~30 samples grouped under `Samples/Layouts`, `Samples/Text`, `Samples/Misc`, `Samples/Developer` — e.g., `OrgChart1`, `Container1`, `TextMarkup1`, `BezierCircle`, `PathAnalysis`, `SendConnectorsToBack`. Picks samples in a `FormSampleRunner` UI and runs them against a live Visio instance. |
| [`VSamples.Docs`](../VisioAutomation_2010/VSamples.Docs/) | Minimal console project (.NET 4.0) holding a small number of canonical examples used in the public documentation. Distinct from `VSamples` in that it is curated for docs, not for exploration. |
| [`DemoIronPython`](../VisioAutomation_2010/DemoIronPython/) | Loose IronPython scripts (`visio.py`, `demo_01_basics.py`, `demo_02_draw_grid.py`, `demo_03_shapesheet.py`) showing the libraries called from Python through CLR interop. Not a csproj. |

---

## 4. Project dependency graph

```
                 ┌──────────────────────┐
                 │   VisioPowerShell    │
                 └──────────┬───────────┘
                            │
                 ┌──────────▼───────────┐
                 │    VisioScripting    │
                 └──────────┬───────────┘
                            │
                 ┌──────────▼───────────┐
                 │ VisioAutomation.Models│
                 └──────────┬───────────┘
                            │
                 ┌──────────▼───────────┐
                 │   VisioAutomation    │
                 └──────────┬───────────┘
                            │
                 ┌──────────▼───────────┐
                 │  Visio 2010 PIA (COM)│
                 └──────────────────────┘

Tests:
  VTest             ──► VisioAutomation
  VTest.Models      ──► VisioAutomation.Models  (and core)
  VTest.Scripting   ──► VisioScripting
  VTest.PowerShell  ──► VisioPowerShell
```

The dependency graph is a strict DAG — there are no upward references. In particular, `VisioAutomation` (core) does **not** reference `VisioAutomation.Models`.

---

## 5. Central concepts

A handful of cross-cutting ideas show up in many places. Each is summarized here with a pointer to the primary code; for definitions of the underlying Visio terms see [GLOSSARY.md](GLOSSARY.md).

### 5.1 ShapeSheet addressing — `Src` and `SidSrc`
Every Visio cell is identified by a `(Section, Row, Cell)` triple — modeled as the [`Src`](../VisioAutomation_2010/VisioAutomation/Core/Src.cs) struct. When the cell is on a specific shape, prefix with the shape ID to get [`SidSrc`](../VisioAutomation_2010/VisioAutomation/Core/SidSrc.cs). These two structs are the low-level vocabulary for every read or write that touches the ShapeSheet.

### 5.2 Batch ShapeSheet I/O — `CellQuery` / `SrcWriter`
COM round trips to Visio are expensive. The library funnels reads through `ShapeSheet.Query.CellQuery` and writes through `ShapeSheet.Writers.SrcWriter` / `SidSrcWriter` so that many cells move per call. Cell-group types like `ShapeFormatCells`, `PageLayoutCells`, `HyperlinkCells` package related cells with strongly-typed properties.

### 5.3 Extension methods over Visio COM types
Rather than wrapping Visio interfaces, the core library extends them. `PageMethods_General`, `PageMethods_Draw`, `PageMethods_Drop`, `PageMethods_ShapeSheet`, and the `ShapeMethods_*` / `MasterMethods_*` families are the public API surface for most page/shape operations.

### 5.4 The DOM — declarative diagram authoring
`VisioAutomation.Models.Dom` defines a serializable tree of `Document` → `Page` → `Shape` you build in memory and then `Render()`. It owns no live COM references until rendering; layouts and OrgChart are built on top of it.

### 5.5 The scripting Client + Target* pattern
`VisioScripting.Client` aggregates command groups; the `Target*` types let callers say *"the active page"* / *"the current selection"* without having to fetch and pass COM objects. `CommandTargetFlags` declares preconditions, validated by `CommandTarget` before the command runs.

### 5.6 Discriminated dispatch — `VisioObjectTarget`
Pages, masters, and shapes all expose ShapeSheet cells, but they are distinct COM types. The internal `VisioObjectTarget` struct + `VisioObjectCategory` enum let helper methods accept any of the three and dispatch correctly without duplicated code.

---

## 6. Packaging

The library is published as the [`VisioAutomation2010`](../NuGet/VisioAutomation2010.nuspec) NuGet package. The nuspec packs the built `VisioAutomation*.dll` and `VisioScripting*.dll` (plus `Microsoft.Msagl.dll` and `GenTreeOps.dll`) under `lib/net40` and declares `Microsoft.Office.Interop.Visio` as a framework assembly reference.

The PowerShell module is shipped separately and installed by copying the build output into the user's PowerShell modules directory; see [`InstallForCurrentUser.ps1`](../VisioAutomation_2010/VisioPowerShell/InstallForCurrentUser.ps1).

---

## 7. Where to start reading

- **Want to see the public surface?** Browse [`VisioAutomation/Extensions/`](../VisioAutomation_2010/VisioAutomation/Extensions/) and [`VisioAutomation/Application/ApplicationHelper.cs`](../VisioAutomation_2010/VisioAutomation/Application/ApplicationHelper.cs).
- **Want to see how high-level scripting is wired?** Read [`Client.cs`](../VisioAutomation_2010/VisioScripting/Client.cs) and one command group, e.g. [`Commands/PageCommands.cs`](../VisioAutomation_2010/VisioScripting/Commands/).
- **Want to understand declarative diagrams?** Read [`VisioAutomation.Models/Dom/Document.cs`](../VisioAutomation_2010/VisioAutomation.Models/Dom/Document.cs), `Page.cs`, `Shape.cs`.
- **Want to see a working example?** Open [`VSamples/Samples/`](../VisioAutomation_2010/VSamples/Samples/) and pick something small.
