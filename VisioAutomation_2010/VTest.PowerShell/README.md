# VTest.PowerShell

Test project for the **VisioPowerShell** module — the cmdlets shipped to the PowerShell Gallery as the `Visio` module.

Current verified counts and environment are recorded in [HANDOVER.md](../../docs/HANDOVER.md). This project's job is to verify cmdlet wiring (parameter binding, pipeline, error handling) inside a real PowerShell runspace, in addition to direct cmdlet smoke tests.

## What it covers

| File | Purpose |
|---|---|
| `BasicTests.cs` | Cmdlet smoke tests run through an in-process PowerShell session. |
| `CmdletBindingTests.cs` | Switch binding, export overwrite, and invalid-point regression tests using the real parameter binder. |
| `ManifestTests.cs` | Compare manifest exports with cmdlet attributes in the built assembly; no Visio required. |
| `SessionTests.cs` | Verify the runspace imports the exact local assembly without module autoloading; no Visio required. |
| `VisioPSSession.cs` | Wrapper around a `System.Management.Automation.Runspaces` runspace plus the registered cmdlets. |
| `Framework/VTestPowerShellSession.cs` | Helpers for invoking cmdlets and consuming their pipeline output. |
| `Framework/VTestPsArray.cs` | Marshaling for cmdlet results that come back as `PSObject[]` / `Array`. |
| `Framework/Extensions/VTestCmdletExtensions.cs` | Convenience extensions on the session for common cmdlet patterns. |

## Test pattern (different from the other test projects)

Unlike `VTest.Models` and `VTest.Scripting`, this project does **not** inherit from `VTest.Framework.VTest`. The shared base class is built around a Visio singleton accessed directly via COM; here, Visio is created and torn down via the cmdlets themselves through a PowerShell session. Different lifecycle, different test surface.

`BasicTests` uses `[ClassInitialize]` to spin up the session and `[ClassCleanup]` to tear it down (closing Visio via the cmdlet, then disposing the runspace). The teardown swallows exceptions deliberately — a teardown failure shouldn't fail an otherwise-green test run.

Register module imports before opening the runspace. `UseCurrentThread` keeps direct helpers and scripted calls on the same thread when they share Visio COM objects. An import after `Open()` does not populate the running session, and execution on a separate runspace thread can fail application-identity checks. Negative binding tests must check the expected error, not merely accept any exception.

## Running

See [docs/BUILDING.md](../../docs/BUILDING.md) and [docs/TESTING.md](../../docs/TESTING.md).
