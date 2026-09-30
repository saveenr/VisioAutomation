# VTest.Scripting

Test project for the **VisioScripting** library — the high-level scripting facade that wraps the lower-level VisioAutomation API into per-subsystem clients (Page, Shape, Text, Document, etc.).

Current verified counts and environment are recorded in [HANDOVER.md](../../docs/HANDOVER.md).

## What it covers

One file per VisioScripting subsystem under test:

| Area | File |
|---|---|
| Application client | `ApplicationTests.cs`, `ClientTests.cs` |
| Document operations | `DocumentTests.cs` |
| Page operations | `PageTests.cs` |
| Shape arrangement | `ArrangeTests.cs`, `GroupTests.cs`, `SelectionTests.cs` |
| Connections | `ConnectTests.cs`, `ConnectionPointTests.cs` |
| Shape attributes | `ControlTests.cs`, `CustomPropTests.cs`, `HyperlinkTests.cs` |
| Text | `ShapeTextTests.cs` |
| ShapeSheet | `ShapeSheetTests.cs` |
| Drop / draw | `DrawManualShapesTests.cs`, `DropContainerTests.cs`, `DropMasterTests.cs` |
| Export | `ExportTests.cs` |
| Developer-mode toggles | `DevTests.cs` |

## Test pattern

Tests inherit from `VTest.Framework.VTest`. Most use `this.GetScriptingClient()` to obtain a `VisioScripting.Client` wired to the shared singleton Visio instance. See [VTest/README.md](../VTest/README.md) for the shared infrastructure.

This project has its own `AssemblyHooks.cs` for `[AssemblyCleanup]` (per-assembly, not inherited).

## Running

See [docs/BUILDING.md](../../docs/BUILDING.md) and [docs/TESTING.md](../../docs/TESTING.md).
