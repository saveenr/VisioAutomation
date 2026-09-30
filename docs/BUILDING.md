# Building, Testing, and Running

Run these commands from the repository root. See [ARCHITECTURE.md](ARCHITECTURE.md) for project structure and [HANDOVER.md](HANDOVER.md) for readiness and release responsibilities.

## Prerequisites

- Windows and Visual Studio 2026 (or its Build Tools) with .NET desktop build tools, matching VisioAutomation.VDX.
- The .NET 10 SDK. The root `global.json` selects the latest installed stable 10.0 feature band, starting at 10.0.100. This is a build-tool dependency, not a runtime retargeting to .NET 10.
- NuGet access for the first restore. All 11 projects are SDK-style and use PackageReference with versions in [Directory.Packages.props](../VisioAutomation_2010/Directory.Packages.props).
- Microsoft Visio for integration tests, samples, and automation. **Compilation does not require Visio**: the Visio 2010 interop assembly comes from NuGet.
- **Windows PowerShell 5.1** for the verified `Visio` automation workflow. Native PowerShell 7 automation and its compatibility shim were not verified in this pass.

Shipping libraries target .NET Framework 4.5.2; tests and VPlayground target 4.7.2. Reference assemblies for both targets are restored from NuGet; separate Developer Pack installations are unnecessary. C# language selection is controlled by `LangVersion` in [Directory.Build.props](../VisioAutomation_2010/Directory.Build.props).

## Build

From PowerShell, locate VS 2026 and restore/build:

```powershell
$vswhere = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe"
$vs = & $vswhere -version '[18.0,19.0)' -products '*' -requires Microsoft.Component.MSBuild -latest -property installationPath
if (-not $vs) { throw 'Visual Studio 2026 MSBuild was not found.' }
$msbuild = Join-Path $vs 'MSBuild\Current\Bin\MSBuild.exe'
& $msbuild VisioAutomation_2010\VisioAutomation2010.slnx -restore -p:Configuration=Debug -m
if ($LASTEXITCODE -ne 0) { throw 'Build failed.' }
```

For shipping artifacts use `-p:Configuration=Release`. Outputs stay in `bin\Debug` or `bin\Release`, without a target-framework suffix. In a VS 2026 Developer Command Prompt, `msbuild` is already on PATH. Opening the SLNX solution in VS 2026 also works.

The canonical solution is `VisioAutomation2010.slnx`. Debug and Release build all eleven projects as Any CPU; legacy platform mappings that skipped projects have been removed. C# 14 is selected explicitly, independently of the .NET Framework runtime targets. The root `NuGet.config` clears machine-local feeds and restores from nuget.org.

## Run all four test projects

The complete suite needs installed, activated Visio and an interactive Windows session. Tests use real COM automation and must run sequentially. Save personal Visio work before starting; tests open and close their own documents and applications.

Using `$vs` from the build command:

```powershell
$vstest = Join-Path $vs 'Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe'
$configuration = 'Debug' # Use Release to validate shipping builds.
$assemblies = 'VTest', 'VTest.Models', 'VTest.Scripting', 'VTest.PowerShell' |
    ForEach-Object { "VisioAutomation_2010\$_\bin\$configuration\$_.dll" }
& $vstest $assemblies /Platform:x64 '/Logger:trx;LogFileName=handover.trx' /ResultsDirectory:TestResults
if ($LASTEXITCODE -ne 0) { throw 'Tests failed. Inspect TestResults\handover.trx.' }
```

Do not add `/Parallel`. Test Explorer also works. Inspect the TRX for failed and skipped tests. Compare Visio processes before and after the run: a completed run should leave no new processes. Interrupted runs may require manual cleanup; do not terminate unrelated Visio sessions.

See [TESTING.md](TESTING.md) for suite design and coverage limitations.

## Continuous integration

[build.yml](../.github/workflows/build.yml) builds Debug and Release on `windows-2025-vs2026` with VS 2026 and the SDK selected in `global.json`, matching VDX. It runs `ManifestTests` and `SessionTests`, which need no Visio installation. The two release-build workflows use the same toolchain and solution. Hosted runners do not have Visio, so a green CI run does not establish that COM integration tests pass. Run the complete suite locally before releasing and retain the TRX evidence.

## Load the PowerShell module

Build Debug, then open a fresh **Windows PowerShell 5.1** session:

```powershell
Import-Module .\VisioAutomation_2010\VisioPowerShell\bin\Debug\Visio.psd1 -Force
Get-Command -Module Visio
```

`LoadFromBinDebug.ps1` performs the development import. Substitute `Release` in the path to verify a Release build. Restart the PowerShell process after rebuilding: removing a module does not unload its .NET assemblies.

For a persistent user installation:

```powershell
& .\VisioAutomation_2010\VisioPowerShell\InstallForCurrentUser.ps1 -Configuration Debug
```

This replaces `Documents\WindowsPowerShell\Modules\Visio`. Close sessions using its DLLs first. Importing directly from the build directory avoids replacing an installed module.

## Produce release artifacts

Build the solution in **Release** first. The NuGet specification packs six runtime DLLs from `VisioScripting\bin\Release` into `lib\net452`, plus the root README:

```powershell
nuget pack NuGet\VisioAutomation2010.nuspec -OutputDirectory TestResults
if ($LASTEXITCODE -ne 0) { throw 'NuGet packaging failed.' }
```

[AcquireNuGetExe.ps1](../NuGet/AcquireNuGetExe.ps1) can acquire `nuget.exe` if needed. Packaging does not publish anything.

Releases have two workflow stages:

| Artifact | Build and create GitHub Release | Publish existing release artifact |
|---|---|---|
| NuGet | [release-nuget.yml](../.github/workflows/release-nuget.yml) | [publish-nuget.yml](../.github/workflows/publish-nuget.yml) |
| PowerShell | [release-psmodule.yml](../.github/workflows/release-psmodule.yml) | [publish-psmodule.yml](../.github/workflows/publish-psmodule.yml) |

Both release flows use Release binaries. Before triggering them, bump the artifact version and roll its `[Unreleased]` changelog entries into a matching `[<version>]` section. Workflows read that **versioned** section. Both stages support `dry_run`; the publish stage requires an existing GitHub Release tag. A release dry run still requires a version whose tag does not already exist.

## Samples

`VSamples` is a WinForms application: select it as the startup project and run a sample against installed Visio. `VSamples.Docs` contains curated documentation examples. `DemoIronPython` holds standalone IronPython examples and loader instructions.
