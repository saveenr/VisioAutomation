# Building, Testing, and Running

Practical notes on building the solution, running the tests, and trying things out locally. For the structure of the projects themselves see [ARCHITECTURE.md](ARCHITECTURE.md).

## Prerequisites

- **Microsoft Visio**, installed locally. The solution targets the Visio 2010 Primary Interop Assembly (`Microsoft.Office.Interop.Visio` v14) but works against newer Visio versions at runtime. Tests and samples instantiate a real Visio process, so Visio must be present on any machine that runs them.
- **Visual Studio 2022** (the .sln declares `VisualStudioVersion = 17.0`). The Build Tools alternative also works. **VS 2026 is not yet supported** — its MSBuild does not resolve targeting packs older than .NET Framework 4.6.2, and most projects target 4.5. Moving to VS 2026 is a Phase 3 item; see [FUTURES.md](FUTURES.md).
- **PowerShell** — required only if you are building/testing/running the `VisioPowerShell` module.

The shipping libraries target .NET Framework 4.5.2, but you do **not** need to install the 4.5.2 Developer Pack — the reference assemblies are supplied by the [`Microsoft.NETFramework.ReferenceAssemblies.net452`](../VisioAutomation_2010/Directory.Packages.props) NuGet package, restored automatically with the rest of the solution. The 4.7.2 reference assemblies (used by the test projects) ship in-box on every supported Windows.

## Building

The exact MSBuild path depends on your VS 2022 install location. From a regular shell (Bash/PowerShell), use the full path:

```sh
# from the repo root, using VS 2022 Community at the default install path
MSBUILD="/c/Program Files/Microsoft Visual Studio/2022/Community/MSBuild/Current/Bin/MSBuild.exe"

# 1. Restore NuGet packages
"$MSBUILD" VisioAutomation_2010/VisioAutomation2010.sln -t:Restore

# 2. Build
"$MSBUILD" VisioAutomation_2010/VisioAutomation2010.sln \
    -p:Configuration=Debug -m
```

From the **Developer Command Prompt for VS 2022** (or Developer PowerShell), `MSBuild.exe` is on PATH and you can drop the full path:

```cmd
msbuild VisioAutomation_2010\VisioAutomation2010.sln -t:Restore
msbuild VisioAutomation_2010\VisioAutomation2010.sln -p:Configuration=Debug -m
```

Or open [`VisioAutomation_2010/VisioAutomation2010.sln`](../VisioAutomation_2010/VisioAutomation2010.sln) in Visual Studio 2022 and build the solution — the IDE handles restore automatically.

The Visio PIA comes from the [`Visio2010.PrimaryInteropAssembly`](../VisioAutomation_2010/Directory.Packages.props) NuGet package, so a clean machine without Visio's developer tools installed will still restore the interop reference. Package versions are centralized in [`Directory.Packages.props`](../VisioAutomation_2010/Directory.Packages.props) (Central Package Management); individual csprojs reference packages without versions.

## Continuous integration

Every push to `master` (and every PR targeting it) is built by [`.github/workflows/build.yml`](../.github/workflows/build.yml) on a GitHub-hosted `windows-latest` runner. The workflow pins MSBuild to VS 2022 (matching local builds) and runs the same restore + build commands documented above.

The CI is **build-only**. Tests need a live Visio install and would require a self-hosted Windows runner; that's planned for Phase 3 alongside automated releases (see [FUTURES.md](FUTURES.md)).

The current build status appears as a badge in the [root README](../readme.md).

## Running the tests

All test projects use **MSTest** and **require a live Visio installation** because they exercise real COM calls.

- From Visual Studio: open Test Explorer, build, run all.
- From the command line:
  ```sh
  vstest.console.exe VisioAutomation_2010\VTest\bin\Debug\VTest.dll
  ```

Tests will launch one or more Visio processes during a run. If a previous run was interrupted, leftover Visio processes can hold file locks — close them before re-running.

| Project | Scope |
|---|---|
| `VTest` | Core library |
| `VTest.Models` | DOM and layouts |
| `VTest.Scripting` | High-level scripting facade |
| `VTest.PowerShell` | PowerShell cmdlets (spins up an in-process PS session) |

## Running the samples

[`VSamples`](../VisioAutomation_2010/VSamples/) is a WinForms exe — set it as the startup project, build, run. The form lists the built-in samples by category; pick one and click run. It will start Visio (if not already running) and execute against a fresh document.

[`VSamples.Docs`](../VisioAutomation_2010/VSamples.Docs/) is a smaller console exe holding the curated examples that appear in the public docs.

## Working with the PowerShell module

For a fast inner loop, build the solution in **Debug** and then in PowerShell:

```powershell
cd VisioAutomation_2010\VisioPowerShell
. .\LoadFromBinDebug.ps1
```

This imports `bin\Debug\Visio.psd1` directly so you can iterate on cmdlets without installing the module.

To install the module for your user (so any PowerShell session can `Import-Module Visio`):

```powershell
cd VisioAutomation_2010\VisioPowerShell
. .\InstallForCurrentUser.ps1
```

The script robocopies the build artifacts to `Documents\WindowsPowerShell\Modules\Visio\`. It will warn if any DLLs are locked by a running PowerShell process — close those sessions first.

## Trying it from IronPython

[`DemoIronPython`](../VisioAutomation_2010/DemoIronPython/) contains stand-alone scripts. The bootstrap loader [`visio.py`](../VisioAutomation_2010/DemoIronPython/visio.py) finds the assemblies (NuGet cache or local build output) and `clr.AddReference`s them. Run e.g. `ipy demo_01_basics.py` with the assemblies on the load path.

## Producing the NuGet package

The package metadata lives in [`NuGet/VisioAutomation2010.nuspec`](../NuGet/VisioAutomation2010.nuspec). It packs the built DLLs from `VisioScripting/bin/debug/` into `lib/net40/` and declares `Microsoft.Office.Interop.Visio` as a framework reference. Build the solution first, then:

```sh
nuget pack NuGet\VisioAutomation2010.nuspec
```

[`NuGet/AcquireNuGetExe.ps1`](../NuGet/AcquireNuGetExe.ps1) helps fetch `nuget.exe` if you don't already have it.

## Known rough edges (cleanup candidates for the 2026 refresh)

See [FUTURES.md](FUTURES.md) for the full backlog and phasing. The build-relevant ones:

- **Mixed target frameworks**: shipping libs are now on .NET 4.5; test projects on .NET 4.7.2. Convergence on a single TFM (4.7.2 across the whole solution) is a Phase 3 item; it also enables moving to VS 2026.
- **`packages.config`** is still in use rather than PackageReference. Modernizing would simplify NuGet handling and CI.
