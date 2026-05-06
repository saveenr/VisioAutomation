# VTest.Models

Test project for the **VisioAutomation.Models** library — DOM, geometry, layout algorithms, and the directed-graph / orgchart drawing models.

45 tests as of 2026-05-04.

## What it covers

| Area | Files |
|---|---|
| DOM (declarative shape construction) | `Dom_Tests.cs`, `Dom_Text_Tests.cs` |
| Layout primitives (boxes, containers) | `Layout_BoxTests.cs`, `Layout_ContainerTests.cs` |
| Geometry / math | `BezierTests.cs`, `BoundingBoxHelperTests.cs` |
| Tree algorithms (used by orgchart) | `Tree_Tests.cs` |
| Drawing scenarios | `DrawModel_DirectedGraph.cs`, `DrawModel_OrgChartTests.cs`, `DrawModel_Grid.cs`, `DrawModel_DataTable.cs` |
| Scripting × Models integration | `Scripting_DropContainerTests.cs` |

## Test pattern

Tests inherit from `VTest.Framework.VTest` (in the `VTest` project) — they use the shared singleton Visio instance and the `GetNewPage()` / `GetScriptingClient()` helpers. See [VTest/README.md](../VTest/README.md) for the shared infrastructure.

This project's own `AssemblyHooks.cs` carries the `[AssemblyCleanup]` that closes Visio at end-of-run (the attribute is per-assembly and not inherited from the base project).

## Visio version sensitivity

`DrawModel_OrgChartTests` and `Dom_DrawOrgChart` opens a stencil whose filename changed in Visio 2013. The test code version-guards on `app.Version` to pick `orgchart.vst` (Visio < 15) vs `orgch_u.vstx` (Visio ≥ 15) — see commit `da9bba0a`. If you add new orgchart-flavored tests, follow the same pattern.

## Running

See [docs/BUILDING.md](../../docs/BUILDING.md) and [docs/TESTING.md](../../docs/TESTING.md).
