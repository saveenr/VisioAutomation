# VTest.Scripting

Test project for the **VisioScripting** library — the high-level scripting facade that wraps the lower-level VisioAutomation API into per-subsystem clients (Page, Shape, Text, Document, etc.).

34 tests as of 2026-05-04.

## What it covers

One file per VisioScripting subsystem under test:

| Area | File |
|---|---|
| Application client | `Scripting_ApplicationTests.cs`, `Scripting_ClientTests.cs` |
| Document operations | `Scripting_DocumentTests.cs` |
| Page operations | `Scripting_PageTests.cs` |
| Shape arrangement | `Scripting_ArrangeTests.cs`, `Scripting_GroupTests.cs`, `Scripting_SelectionTests.cs` |
| Connections | `Scripting_ConnectTests.cs`, `Scripting_ConnectionPointTests.cs` |
| Shape attributes | `Scripting_ControlTests.cs`, `Scripting_CustomPropTests.cs`, `Scripting_HyperlinkTests.cs` |
| Text | `Scripting_ShapeText_Tests.cs` |
| ShapeSheet | `Scripting_ShapeSheetTests.cs` |
| Drop / draw | `Scripting_DrawManualShapes.cs`, `Scripting_DropContainerTests.cs`, `Scripting_DropMasterTests.cs` |
| Export | `Scripting_ExportTests.cs` |
| Developer-mode toggles | `Scripting_DevTests.cs` |

## Test pattern

Tests inherit from `VTest.Framework.VTest`. Most use `this.GetScriptingClient()` to obtain a `VisioScripting.Client` wired to the shared singleton Visio instance. See [VTest/README.md](../VTest/README.md) for the shared infrastructure.

This project has its own `AssemblyHooks.cs` for `[AssemblyCleanup]` (per-assembly, not inherited).

## Running

See [docs/BUILDING.md](../../docs/BUILDING.md) and [docs/TESTING.md](../../docs/TESTING.md).
