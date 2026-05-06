# Changelog — Visio PowerShell module

All notable changes to the [`Visio`](https://www.powershellgallery.com/packages/Visio) PowerShell module are documented here.

For the bundled .NET library's release history see [`NuGet/CHANGELOG.md`](../../NuGet/CHANGELOG.md).

The format follows [Keep a Changelog 1.1.0](https://keepachangelog.com/en/1.1.0/) and the module follows [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

> **Contributors:** when a change affects what consumers of this module see (cmdlets added/changed/removed, parameter changes, behavior changes, minimum supported PowerShell or .NET Framework version), add an entry under `[Unreleased]`. Pure internal/build/docs changes don't need to appear here.

## [Unreleased]

### Fixed
- **`Get-VisioShape`** now declares an explicit `DefaultParameterSetName = "shapebyname"`. Previously the cmdlet had three parameter sets (`active`, `shapebyname`, `shapebyid`) but no default, so a no-args `Get-VisioShape` call relied on PowerShell nondeterministically picking a set; under stricter PowerShell configurations it could throw `AmbiguousParameterSet`. The "no args returns every shape on the page" behavior is now an explicit, documented part of the cmdlet rather than an accidental fallthrough. Closes [#130](https://github.com/saveenr/VisioAutomation/issues/130).
- **`Get-VisioLockCells`** now calls `WriteObject(dic)` instead of `WriteObject(dic, true)`, matching its three sibling "Get a dictionary keyed by shape" cmdlets (`Get-VisioCustomProperty`, `Get-VisioHyperlink`, `Get-VisioUserDefinedCell`). Pure consistency fix: PowerShell special-cases `IDictionary` and doesn't enumerate it across the pipeline regardless of the flag, so observable behavior is unchanged. Closes [#129](https://github.com/saveenr/VisioAutomation/issues/129).

## [4.6.1] - 2026-05-03

First release cut from the `2026_Refresh` work. Bundles the Phase 1 cleanup work and four cmdlet bug fixes.

### Fixed
- **`Lock-VisioShape` / `Unlock-VisioShape`** — the 20 lock-flag switches (`-Aspect`, `-Width`, `-Height`, `-MoveX`, `-MoveY`, `-Delete`, `-Format`, `-Rotate`, etc.) now actually bind. Previously the switches were declared without `[Parameter]` attributes, so PowerShell silently ignored them and both cmdlets were no-ops regardless of the flags passed.
- **`Export-VisioShape`** — the file-existence check was inverted. Previously, exporting to a fresh path without `-Overwrite` raised *"File already exists"*, while writing to an existing path silently overwrote regardless of `-Overwrite`. Now the cmdlet writes fresh paths normally and refuses to overwrite existing files unless `-Overwrite` is passed.
- **`New-VisioShape`** — the polyline-≥2-points and Bezier-≥4-points validations now actually throw `ArgumentOutOfRangeException`. Previously they constructed the exception object without throwing it, leaving invalid input to fail later inside Visio.

### Changed
- **Minimum .NET Framework runtime is now 4.5.2** (was 4.5). The bundled DLLs target 4.5.2; consumers running .NET Framework 4.5 or 4.5.1 will need to install 4.5.2 (universally available on supported Windows).

## Earlier versions

Versions 4.6.0 and earlier predate this changelog. See the [git history](https://github.com/saveenr/VisioAutomation/commits/master/) and [release tags](https://github.com/saveenr/VisioAutomation/releases) for details.
