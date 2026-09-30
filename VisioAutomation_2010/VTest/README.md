# VTest

Test project for the **VisioAutomation** core library, and home of the shared Visio lifecycle infrastructure used by `VTest.Models` and `VTest.Scripting`. `VTest.PowerShell` manages its own cmdlet session lifecycle.

The largest of the four test projects. Current verified counts and environment are recorded in [HANDOVER.md](../../docs/HANDOVER.md).

## What it covers

| Area | Files |
|---|---|
| Core types | `Core/CellValueLiteralTests.cs`, `Core/TypeTests.cs`, `Core/PageHelperTests.cs`, `Core/ConnectionPointTests.cs` |
| Application-level | `Core/Application/ApplicationHelperTests.cs`, `Core/Application/XmlErrorLogTests.cs` |
| Cell records | `Core/CellRecords/CellRecordTests.cs` |
| Shape APIs | `Core/Shapes/*.cs` (Connector, Hyperlink, Geometry, CustomProperties, UserDefinedCells, Control, ShapeHelper) |
| ShapeSheet read/write | `Core/ShapeSheet/ShapeSheetWriterTests.cs`, `Core/ShapeSheet/ShapeSheetQueryTests.cs` |
| Text formatting | `Core/Text/TextFormatTests.cs`, `Core/Text/TextUtilTests.cs` |
| Extension methods | `Core/Extensions/*.cs` (Application, Document, Page, Selection, etc.) |
| Connectivity analyzers | `Analyzers/ConnectionAnalysisTests.cs`, `Analyzers/PathTests.cs`, `Analyzers/ConnectivityMap.cs` |
| Misc utilities | `Utilities/ArraySegmentTests.cs` |

## Shared infrastructure (in `Framework/`)

The other test projects reference these via project reference to `VTest`:

- **`VTest.cs`** — base class for nearly every test class in this suite. Provides `GetVisioApplication()`, `GetNewPage()`, `GetNewDoc()`, `GetScriptingClient()` and friends. Marked `[TestClass]` itself but contains no test methods.
- **`VTestAppRef.cs`** — per-testhost Visio singleton. Recreates the instance on `COMException` (i.e., if Visio closed externally). `QuitVisioApplication()` is the cleanup entrypoint.
- **`AssemblyHooks.cs`** — this project's `[AssemblyCleanup]` calling `VTestAppRef.QuitVisioApplication()`. Each test project has its own copy because `[AssemblyCleanup]` is per-assembly and not inherited.
- **`AssertUtil.cs`**, **`VTestExtensions.cs`**, **`VTestGlobals.cs`**, **`VTestHelper.cs`** — shared assertions and helpers.
- **`VTestScriptingClient.cs`** — `VisioScripting.Client` wrapper that captures debug/verbose/user output to a useful place during tests.

## Data files

`datafiles/` contains real Visio XML and document fixtures (`XMLErrorLog_*.txt`, `directed_graph_*.xml`, `orgchart_1.xml`, `template_router.vdx`, `vdx_with_warnings_1.vdx`). All marked `CopyToOutputDirectory=Always`. **Don't** add `[DeploymentItem]` attributes — commit `5cbf11cd` removed those because they triggered VS Test Explorer's deployment mode and dropped runtime dependencies on the floor.

## Running

See [docs/BUILDING.md](../../docs/BUILDING.md) for the basics; [docs/TESTING.md](../../docs/TESTING.md) for design and gotchas.
