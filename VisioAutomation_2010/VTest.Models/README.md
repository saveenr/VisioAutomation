# VTest.Models

Test project for the **VisioAutomation.Models** library — DOM, geometry, layout algorithms, and the directed-graph / orgchart drawing models.

45 tests as of 2026-05-04.

## What it covers

| Area | Files |
|---|---|
| DOM (declarative shape construction) | `DomTests.cs`, `DomTextTests.cs` |
| Layout primitives (boxes, containers) | `BoxLayoutTests.cs`, `ContainerLayoutTests.cs` |
| Geometry / math | `BezierTests.cs`, `BoundingBoxHelperTests.cs` |
| Tree algorithms (used by orgchart) | `TreeTests.cs` |
| Drawing scenarios | `DirectedGraphDrawModelTests.cs`, `OrgChartDrawModelTests.cs`, `GridDrawModelTests.cs`, `DataTableDrawModelTests.cs` |
| Scripting × Models integration | `DropContainerScriptingTests.cs` |

## Test pattern

Tests inherit from `VTest.Framework.VTest` (in the `VTest` project) — they use the shared singleton Visio instance and the `GetNewPage()` / `GetScriptingClient()` helpers. See [VTest/README.md](../VTest/README.md) for the shared infrastructure.

This project's own `AssemblyHooks.cs` carries the `[AssemblyCleanup]` that closes Visio at end-of-run (the attribute is per-assembly and not inherited from the base project).

## Visio version sensitivity

`OrgChartDrawModelTests` and `DomTests.DrawOrgChart` open a stencil whose filename changed in Visio 2013. The test code version-guards on `app.Version` to pick `orgchart.vst` (Visio < 15) vs `orgch_u.vstx` (Visio ≥ 15) — see commit `da9bba0a`. If you add new orgchart-flavored tests, follow the same pattern.

## Running

See [docs/BUILDING.md](../../docs/BUILDING.md) and [docs/TESTING.md](../../docs/TESTING.md).
