# Findings: moving the Visio PowerShell module to PowerShell 7

Experiment [#277](https://github.com/saveenr/VisioAutomation/issues/277), run 2026-11-20 on a local branch (`experiment/powershell7`, never pushed). This note is the deliverable; the experiment code is throwaway. It feeds the spike [#249](https://github.com/saveenr/VisioAutomation/issues/249) and the migration [#175](https://github.com/saveenr/VisioAutomation/issues/175).

**Setup:** Windows 11, PowerShell 7.6.10 (installed; .NET 10.0.15), PowerShell 7.7.0 (portable zip, hash checked; .NET 11.0.0), Windows PowerShell 5.1.26100, .NET SDKs 10.0.404 and 11.0.100, live Visio Professional. Module 5.1.0 source. The first pass used 7.6.10 only; 7.7.0 and the `net11` build were added the same day after 7.7.0 turned out to be out (the first version of this note wrongly said PowerShell 7 runs on .NET 10).

## Summary

The move is much smaller than expected. The libraries and the cmdlet project compile for `net10.0-windows` **and for `net11.0-windows`, with only project-file edits and no source-code changes**, and the resulting module loaded in PowerShell 7 and ran the cmdlets we tried against live Visio, including a directed-graph draw that goes through MSAGL. Even the **unchanged net472 build imports and runs in PowerShell 7**. The existing test suites also ported with one source change and all 360 tests passed on net11 against live Visio. The real work is in packaging, the .NET version choice (it decides which PowerShell 7 versions are supported, see problem 1) and keeping Windows PowerShell 5.1 supported.

## What happened, step by step

1. **Today's net472 module in PowerShell 7, unchanged: works.** `Import-Module .\Visio.psd1` succeeds in `pwsh` and exports all 64 cmdlets. The net472 assemblies run on .NET 10 through its .NET Framework compatibility. `New-VisioApplication`, `New-VisioDocument`, `Open-VisioDocument`, `Set-VisioText`, `Export-VisioPage` and `Close-VisioApplication` all worked. Caveat: this is an accident of compatibility, not something to ship. Windows PowerShell 5.1 and PowerShell 7 disagree about what the module is (the manifest has no `CompatiblePSEditions`), and the net472 DLLs run in a compatibility mode that can break on any API Microsoft removed.
2. **A `net10.0-windows` build: compiles with zero source errors.** Only the four project files changed (see "Changes that were needed"). Warnings: NU1701 for three packages restored through the .NET Framework fallback (see below).
3. **The net10 module in PowerShell 7.6 and 7.7: works in both.** Same tour as step 1, plus `Import-VisioModel` piped to `Out-VisioApplication` for a directed graph (11 shapes drawn). MSAGL runs even though its package is a .NET Framework one.
4. **The net10 module in Windows PowerShell 5.1: does not load.** `Could not load file or assembly 'System.Runtime, Version=10.0.0.0'`. One build cannot serve both hosts, so supporting both means either two builds (multi-targeting) or two modules.
5. **A `net11.0-windows` build (needs SDK 11 via `global.json`): compiles with zero errors and the same project-file edits.** It loads and runs in PowerShell 7.7.0 (same tour, graph drew 11 shapes) and fails to load in 7.6.10 and 5.1 (problem 1). New warnings versus net10: 60 `WFO2000` (a WinForms analyzer note about getters that return new instances, for example `Text/FieldConstants.cs`; style noise, not a break) and 2 `SYSLIB0051` (the obsolete serialization constructor in `Exceptions/AutomationException.cs`; fix before it becomes an error in a later .NET).
6. **No orphan Visio processes.** Every `New-VisioApplication` in the tours was closed by `Close-VisioApplication`; the only Visio process left was one that existed before the experiment.

## Changes that were needed for the net10 build

All in project files; none in C# source.

- **TargetFramework** to `net10.0-windows` on `VisioAutomation`, `VisioAutomation.Models`, `VisioScripting` and `VisioPowerShell`.
- **Remove the legacy `<Reference Include="System" />` items** (`System`, `System.Core`, `System.Drawing`, `System.Data`, `System.Xml`, `Microsoft.CSharp` and so on). On modern .NET they do not resolve (warnings MSB3243 and MSB3245) and the framework supplies them anyway.
- **`<UseWindowsForms>true</UseWindowsForms>` on `VisioAutomation`.** Without it: `error CS0234: The type or namespace name 'Forms' does not exist in the namespace 'System.Windows'` in `UserInterfaceHelper.cs`. WinForms and `System.Drawing` types are used in four files: `Application/UserInterfaceHelper.cs`, `Extensions/WindowMethods.cs`, `Color/ColorRgb.cs` and `ApplicationCommands.cs`.
- **Drop the `Microsoft.NETFramework.ReferenceAssemblies.net472` package** (not needed off net472).
- **Replace `Microsoft.PowerShell.3.ReferenceAssemblies` with `PowerShellStandard.Library` 5.1.1** in `VisioPowerShell` (and add its version to `Directory.Packages.props`). The compile worked with no code changes.
- **`<CopyLocalLockFileAssemblies>true</CopyLocalLockFileAssemblies>` on `VisioPowerShell`.** Without it the output folder lacks `Microsoft.Msagl.dll`, `GenTreeOps.dll` and the Visio PIA, and `Import-Module` fails with `Could not load file or assembly 'Microsoft.Msagl'`. SDK-style class libraries on modern .NET do not copy NuGet dependencies by default; net472 did.

## Problems and risks found

1. **A module can only load in a PowerShell 7 that runs on the same or a newer .NET.** PowerShell 7.6.10 runs on **.NET 10.0.15**; PowerShell 7.7.0 (released 2026-11-10) runs on **.NET 11.0.0**. Measured, with the same module built both ways:

   | Module build | 7.6.10 (.NET 10) | 7.7.0 (.NET 11) | Windows PowerShell 5.1 |
   |---|---|---|---|
   | `net10.0-windows` | loads, runs | loads, runs | fails (`System.Runtime 10.0.0.0`) |
   | `net11.0-windows` | **fails** (`System.Runtime 11.0.0.0` not found) | loads, runs | fails |

   So a `net11` target is fine for 7.7 and later and is a plain error on 7.6, 7.5 and 7.4. A `net10` target covers 7.6 and 7.7. The target therefore decides which PowerShell 7 versions are supported, and the direction in [#175](https://github.com/saveenr/VisioAutomation/issues/175) (".NET 11") means "PowerShell 7.7 and later only". The failure mode is unfriendly: in 7.6, `Import-Module` of the `net11` build prints the assembly error and leaves **no cmdlets loaded** (0 exported). A manifest `PowerShellVersion` of `7.7` would turn that into a clear message. Separately, the repo's `global.json` pins SDK 10, so a `net11` target does not build until it is changed to 11 (`NETSDK1045` otherwise).
2. **Three dependencies restore only through the .NET Framework fallback (NU1701):** the Visio PIA (`Visio2010.PrimaryInteropAssembly`), `Microsoft.Automatic.Graph.Layout` (MSAGL) and `GenTreeOps`. All three worked at runtime in the tours, but the warning means they are not built for modern .NET and may break on a specific code path we did not exercise. The PIA is the one that matters most: it is COM interop, so test it on every call path (this is what [#249](https://github.com/saveenr/VisioAutomation/issues/249) asks). The experiment did not test events, `Selection`, UI and window methods, or the MSAGL options beyond one graph.
3. **A stray `System.Management.Automation.dll` lands in the module folder.** `PowerShellStandard.Library` is copy-local, so a 5.1 reference assembly is copied next to the module. PowerShell 7 ignored it and loaded its own (7.6.0.500 from `C:\Program Files\PowerShell\7`), but it should be excluded (`PrivateAssets` or `ExcludeAssets`) before anything ships. The same file already appears in today's net472 output.
4. **One build cannot serve both PowerShell hosts** (step 4 above). If Windows PowerShell 5.1 users must keep working, ship net472 and `net10.0-windows` builds: either one module with a per-edition folder layout, or two modules. The repo sets `AppendTargetFrameworkToOutputPath=false`, so a multi-target build would overwrite one TFM's output with the other's; that property has to go (or the output path become per-TFM), and the scripts below must follow.
5. **The manifest needs work.** `Visio.psd1` says `PowerShellVersion = '5.1'` and has no `CompatiblePSEditions`. It still imported in PowerShell 7, but a real PS 7 module needs `CompatiblePSEditions`, the right `PowerShellVersion`, and a decision about `RequiredAssemblies` (the list names `Microsoft.Msagl.dll` and `GenTreeOps.dll`, which have to be present).
6. **Packaging and CI assume the current layout.** `release-psmodule.yml` stages `VisioPowerShell\bin\Release` and checks for four files; `InstallForCurrentUser.ps1` globs `bin\<config>\*.dll`; `scripts\Test-AssemblyMetadata.ps1` runs under Windows PowerShell 5.1 (`shell: powershell`) and reads the built assemblies. All need review for a layout with more than one target. The `msbuild` call on the solution and the `windows-2025-vs2026` runner should handle `net10.0-windows`, but this was not run in CI.
7. **The test suites port with one source change and pass: 360 of 360 on net11.** The four test projects were retargeted to `net11.0-windows` and run with `dotnet test` against live Visio: `VTest` 108, `VTest.Models` 118, `VTest.Scripting` 46, `VTest.PowerShell` 88. That equals the net472 baseline count of 360, with no failures and no skips. The `VTest.PowerShell` runspace host ran on the real PowerShell 7.7.0 hosting package. Details, all found by doing it:
   - **Project-file changes per test project:** target `net11.0-windows`; `PlatformTarget` x64 (the old runs used `vstest.console /Platform:x64`); remove the legacy `System.*` references, the binding-redirect properties and the net472 reference-assemblies package; replace `Microsoft.PowerShell.3.ReferenceAssemblies` with `Microsoft.PowerShell.SDK` 7.7.0 (the version must match the PowerShell 7 / .NET line being tested); add `Microsoft.NET.Test.Sdk`. I also set `UseWindowsForms` on all four; I did not check which of them actually need it.
   - **One source change:** `VTest.PowerShell/Framework/VTestPowerShellSession.cs` used `System.Management.Automation.RunspaceInvoke`, which does not exist in PowerShell 7's SDK (`CS0234`). The field was created and disposed but never used, so it was deleted.
   - **`Microsoft.NET.Test.Sdk` 18.12.0 failed to restore** (`Microsoft.TestPlatform.ObjectModel` 18.12.0 not yet on nuget.org, `NU1102`); 18.10.1 worked. Check the version at the time.
   - **The test command changes.** `docs/BUILDING.md` runs `vstest.console.exe ... /Platform:x64` from Visual Studio; on net11 use `dotnet test` (which also needs the SDK 11 pin).
   - **The solution no longer builds as a whole.** `VSamples`, `VSamples.Docs` and `VPlayground` are net472 executables that reference the libraries, so they fail restore (`NU1201`) once the libraries are net11. They have to be retargeted too, or kept out of the solution for the new line. `VSamples.Docs` matters most because it is the docs sample check ([`decisions/doc-sample-checking.md`](../decisions/doc-sample-checking.md)).
   - **No orphan Visio processes** after all four suites; the existing `[AssemblyCleanup]` hooks held on the new runtime. The per-testhost Visio singleton worked unchanged.
8. **Existing `Get-VisioApplication` behavior is unchanged** (returns nothing when only a manually started Visio is running), same in 5.1 and 7. Not a migration issue; noted because it confused the first probe.

## Control: failures that were not PowerShell 7's fault

`New-VisioShape -Master $m -Position ...` failed with `Unresolved Target` in my tour, and `Get-VisioShapeCells -Cells` was a bad parameter. Both were my usage; the identical script gave the same result in Windows PowerShell 5.1 on the unmodified module. Lesson for the real migration: run every check in 5.1 as well and compare, so usage mistakes are not mistaken for porting problems.

## Not tested

- CI, `Publish-VisioPSToGallery.ps1` and `InstallForCurrentUser.ps1`. The suites were run once, in Debug, on net11 with the PowerShell 7.7.0 SDK; not on net10 / PowerShell 7.6, not in Release, and not with both targets built from one checkout.
- Most of the 64 cmdlets. The tours covered about ten, plus the directed-graph path.
- Visio versions other than the one installed, 32-bit Visio, and Windows PowerShell 5.1 on the net10 build beyond the load failure.
- NetOffice and replacing the PIA ([#158](https://github.com/saveenr/VisioAutomation/issues/158)); out of scope.

## Suggested order for the real migration

1. Decide the target .NET version (problem 1: `net11` means PowerShell 7.7+ only, `net10` covers 7.6 and 7.7) and whether 5.1 stays supported (problem 4). These two decisions shape everything else.
2. Multi-target the four libraries, with the project-file changes above, and fix the output-path property.
3. Test the PIA, MSAGL and WinForms-touching code paths on the new target against live Visio; this is the real risk.
4. Port the test projects (problem 7: mostly project-file edits and one deleted field); run the full suite on each target, and retarget `VSamples`, `VSamples.Docs` and `VPlayground` so the solution builds.
5. Update the manifest, packaging scripts, `Test-AssemblyMetadata.ps1` and the two release workflows, then do the dry runs.
6. Document the PowerShell 7 support and the version-compatibility tables.

## Reproducing

The experiment branch is local only. Its changes were the four csproj edits and the one `Directory.Packages.props` line above (plus `global.json` set to SDK 11.0.100 for the `net11` build); applying them to a clean checkout and running `dotnet build VisioAutomation_2010\VisioPowerShell\VisioPowerShell.csproj -t:Restore,Build` reproduces the module in `VisioPowerShell\bin\Debug` for whichever target the csprojs name. PowerShell 7.7.0 was run from the official portable zip (`PowerShell-7.7.0-win-x64.zip`, SHA256 checked against the release's `hashes.sha256`), side by side with the installed 7.6.
