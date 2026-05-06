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

## Quirks

- **Hardcoded `System.Management.Automation` reference.** This project's csproj references the PSv3 SMA assembly via an absolute GAC path (`C:\Windows\Microsoft.NET\assembly\GAC_MSIL\System.Management.Automation\v4.0_3.0.0.0__31bf3856ad364e35\...`). `VTest` has the same problem to a different path. Phase 3 SDK migration's Pass 1 swaps both to the `Microsoft.PowerShell.3.ReferenceAssemblies` package.
- **`MSB3270` warning.** Every build of this project emits a "processor architecture mismatch" warning because it's `AnyCPU` while `VTest` (which it references) is pinned to `x86`. Tracked in [docs/FUTURES.md](../../docs/FUTURES.md) under *General cleanup of the test projects*; resolved as part of Phase 3 SDK migration's Pass 2b.

## Running

See [docs/BUILDING.md](../../docs/BUILDING.md) and [docs/TESTING.md](../../docs/TESTING.md).
