# VTest.PowerShell

Test project for the **VisioPowerShell** module — the cmdlets shipped to the PowerShell Gallery as the `Visio` module.

4 tests as of 2026-05-04 (the smallest of the four test projects). The cmdlet surface is large but mostly thin wrappers over the same library code that `VTest.Scripting` already exercises end-to-end; this project's job is to verify the cmdlet wiring (parameter binding, pipeline, error handling) actually works inside a real PowerShell runspace.

## What it covers

| File | Purpose |
|---|---|
| `VisioPS_Basic_Tests.cs` | Cmdlet smoke tests run through an in-process PowerShell session. |
| `VisioPS_Session.cs` | Wrapper around a `System.Management.Automation.Runspaces` runspace plus the registered cmdlets. |
| `Framework/VTestPowerShellSession.cs` | Helpers for invoking cmdlets and consuming their pipeline output. |
| `Framework/VTestPsArray.cs` | Marshaling for cmdlet results that come back as `PSObject[]` / `Array`. |
| `Framework/Extensions/VTestCmdletExtensions.cs` | Convenience extensions on the session for common cmdlet patterns. |

## Test pattern (different from the other test projects)

Unlike `VTest.Models` and `VTest.Scripting`, this project does **not** inherit from `VTest.Framework.VTest`. The shared base class is built around a Visio singleton accessed directly via COM; here, Visio is created and torn down via the cmdlets themselves through a PowerShell session. Different lifecycle, different test surface.

`VisioPS_Basic_Tests` uses `[ClassInitialize]` to spin up the session and `[ClassCleanup]` to tear it down (closing Visio via the cmdlet, then disposing the runspace). The teardown swallows exceptions deliberately — a teardown failure shouldn't fail an otherwise-green test run.

## Running

See [docs/BUILDING.md](../../docs/BUILDING.md) and [docs/TESTING.md](../../docs/TESTING.md).
