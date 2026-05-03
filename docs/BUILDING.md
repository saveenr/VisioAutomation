# Building, Testing, and Running

Practical notes on building the solution, running the tests, and trying things out locally. For the structure of the projects themselves see [ARCHITECTURE.md](ARCHITECTURE.md).

## Prerequisites

- **Microsoft Visio**, installed locally. The solution targets the Visio 2010 Primary Interop Assembly (`Microsoft.Office.Interop.Visio` v14) but works against newer Visio versions at runtime. Tests and samples instantiate a real Visio process, so Visio must be present on any machine that runs them.
- **Visual Studio 2022** (the .sln declares `VisualStudioVersion = 17.0`). The Build Tools alternative also works.
- **.NET Framework developer packs**: 4.0, 4.5, and 4.7.2 — the projects target a mix of these (see the table in [ARCHITECTURE.md §3.2](ARCHITECTURE.md#32-test-projects)).
- **PowerShell** — required only if you are building/testing/running the `VisioPowerShell` module.
- **NuGet** restore (Visual Studio handles this automatically; `nuget restore` from the command line works too).

## Building

```sh
# from the repo root
msbuild VisioAutomation_2010\VisioAutomation2010.sln /restore /p:Configuration=Debug
```

Or open [`VisioAutomation_2010/VisioAutomation2010.sln`](../VisioAutomation_2010/VisioAutomation2010.sln) in Visual Studio and build the solution.

The Visio PIA comes from the [`Visio2010.PrimaryInteropAssembly`](../VisioAutomation_2010/VisioAutomation/packages.config) NuGet package, so a clean machine without Visio's developer tools installed will still restore the interop reference.

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

- **Mixed target frameworks**: production projects target .NET 4.0; test projects target a mix of .NET 4.5 and .NET 4.7.2. Worth consolidating.
- **MSTest is on a beta**: `MSTest.TestFramework` 2.0.0-beta2. Either pin to a current stable, or migrate to a newer test framework.
- **`packages.config`** is still in use rather than PackageReference. Modernizing would simplify NuGet handling and CI.
- **`DownloadFromPowerShellGallery.ps1`** is mis-named — it currently loads locally from `bin\Debug` rather than fetching from the PowerShell Gallery.
- **No CI configuration** in the repo today. A simple GitHub Actions workflow that at least builds the solution would catch breakage early. (Tests need Visio, so they would have to run on a self-hosted Windows runner with Visio installed.)
