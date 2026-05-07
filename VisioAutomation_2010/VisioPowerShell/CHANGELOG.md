# Changelog — Visio PowerShell module

All notable changes to the [`Visio`](https://www.powershellgallery.com/packages/Visio) PowerShell module are documented here.

For the bundled .NET library's release history see [`NuGet/CHANGELOG.md`](../../NuGet/CHANGELOG.md).

The format follows [Keep a Changelog 1.1.0](https://keepachangelog.com/en/1.1.0/) and the module follows [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

> **Contributors:** when a change affects what consumers of this module see (cmdlets added/changed/removed, parameter changes, behavior changes, minimum supported PowerShell or .NET Framework version), add an entry under `[Unreleased]`. Pure internal/build/docs changes don't need to appear here.

## [Unreleased]

### Fixed
- **`Set-VisioUserDefinedCell`** &mdash; `-Value` and `-Prompt` parameters now encode their string arguments as Visio formulas before passing them to the underlying `UserDefinedCellHelper.Set`. Regression introduced by the #144 cross-layer migration in [4.7.0]: `Set-VisioUserDefinedCell -Name 'X' -Value 'foo'` would surface an `ArgumentException` ("Visio rejected the formula ... use SetString ... see #144") from the .NET-side detect-and-rethrow, because &mdash; unlike the parallel `Set-VisioCustomProperty` cmdlet, where `VisioScripting.CustomPropertyCommands` calls `EncodeValues()` as a backstop &mdash; the UDC scripting layer has no equivalent backstop. Cmdlet now uses the new typed setter (`SetString`) for `-Value` and explicit `Core.CellValue.EncodeValue` for `-Prompt`. New regression test in `VTest.PowerShell` covers the path going forward.

## [4.7.0] - 2026-05-06

Headline change: typed setters on `CustomPropertyCells` / `UserDefinedCellCells` plus a friendly diagnostic on bad formulas, closing the long-running thread that started with [#117](https://github.com/saveenr/VisioAutomation/issues/117). Plus an audit-pass making `Get-*` cmdlets' positional bindings consistent across the module.

### Added
- **Typed setters on `CustomPropertyCells` and `UserDefinedCellCells`** for setting cell values without having to think about Visio's formula encoding. Each setter writes a correctly-encoded Visio formula and (where applicable) sets the `Type` cell to match. From PowerShell:

  ```powershell
  $cp = New-Object VisioAutomation.Shapes.CustomPropertyCells
  $cp.SetString("hello")        # Type=0 (String)
  $cp.SetNumber(42)             # Type=2 (Number)
  $cp.SetBool($true)            # Type=3 (Boolean)
  $cp.SetDate([datetime]::Now)  # Type=5 (Date)
  $cp.SetFormula("=...")        # raw escape hatch
  ```

  `UserDefinedCellCells` exposes `SetString` and `SetFormula`. The setters become the recommended replacement for raw `$cells.Formula = ...` assignment. Closes [#144](https://github.com/saveenr/VisioAutomation/issues/144).
- **`Formula` property on `CustomPropertyCells` and `UserDefinedCellCells`** as the canonical name (renamed from `Value` to surface that the cell stores a Visio formula, not a literal value).

### Deprecated
- **`CustomPropertyCells.Value` and `UserDefinedCellCells.Value`** are now `[Obsolete]` aliases for `Formula`. Existing PowerShell scripts that read or write `$cells.Value` keep working unchanged through the deprecation window. Migration: rename `$cells.Value` to `$cells.Formula`, or use the new typed setters. Part of [#144](https://github.com/saveenr/VisioAutomation/issues/144).

### Fixed
- **`Get-VisioShape`** now declares an explicit `DefaultParameterSetName = "shapebyname"`. Previously the cmdlet had three parameter sets (`active`, `shapebyname`, `shapebyid`) but no default, so a no-args `Get-VisioShape` call relied on PowerShell nondeterministically picking a set; under stricter PowerShell configurations it could throw `AmbiguousParameterSet`. The "no args returns every shape on the page" behavior is now an explicit, documented part of the cmdlet rather than an accidental fallthrough. Closes [#130](https://github.com/saveenr/VisioAutomation/issues/130).
- **`Get-VisioLockCells`** now calls `WriteObject(dic)` instead of `WriteObject(dic, true)`, matching its three sibling "Get a dictionary keyed by shape" cmdlets (`Get-VisioCustomProperty`, `Get-VisioHyperlink`, `Get-VisioUserDefinedCell`). Pure consistency fix: PowerShell special-cases `IDictionary` and doesn't enumerate it across the pipeline regardless of the flag, so observable behavior is unchanged. Closes [#129](https://github.com/saveenr/VisioAutomation/issues/129).

### Changed
- **`Set-VisioCustomProperty`** &mdash; when callers pass a manually-constructed `CustomPropertyCells` via `-Cells` whose `Formula` (formerly `Value`) field is set to a raw string instead of an encoded Visio formula, the cmdlet now surfaces an `ArgumentException` with a self-explanatory message pointing at the new typed setters (`SetString` / `SetNumber` / `SetBool` / `SetDate`) and `EncodeValues()`. Previously this path raised an opaque `COMException: #NAME?` from the underlying COM call. The default `Set-VisioCustomProperty -Value "x"` flow is unaffected (the cmdlet pre-encodes internally). Part of [#144](https://github.com/saveenr/VisioAutomation/issues/144).
- **Get-* cmdlet positional parameters &mdash; full audit pass.** Eleven `Get-*` cmdlets gain consistent positional bindings so the natural shorthand forms (`Get-VisioPage "Page-1" $doc`, `Get-VisioCustomProperty $shape`, etc.) work as users intuit. Closes [#143](https://github.com/saveenr/VisioAutomation/issues/143) (and supersedes the narrow [#142](https://github.com/saveenr/VisioAutomation/issues/142) entry below). The convention adopted is: cmdlets with both `-Name` and a single object context have `-Name` at position 0 and the object (`-Document` / `-Page`) at position 1; cmdlets with just an object context have it at position 0.
  - `Get-VisioPage`: `-Document` at position 1, `-ID` at position 0 (in its `pagebyid` set), explicit `DefaultParameterSetName = "pagebyname"` to make the no-args case deterministic (same fix shape as [#130](https://github.com/saveenr/VisioAutomation/issues/130) on `Get-VisioShape`).
  - `Get-VisioShape`: `-Name` and `-ID` at position 0 (each in its own set), `-Page` at position 1.
  - `Get-VisioDocument`: `-Name` at position 0, explicit `DefaultParameterSetName = "docbyname"`.
  - `Get-VisioCustomProperty`, `Get-VisioHyperlink`, `Get-VisioLockCells`, `Get-VisioControl`, `Get-VisioUserDefinedCell`, `Get-VisioText`, `Get-VisioShapeCells`: `-Shape` at position 0.
  - `Get-VisioPageCells`: `-Page` at position 0.
- **`Get-VisioMaster`** &mdash; `-Document` parameter is now positional at `Position = 1`, so the natural form `Get-VisioMaster "Group" $doc` works as users intuit. Previously only `-Name` had a position; an unnamed second positional argument either errored or got string-coerced into `-Name`'s array. The bind change is back-compat-safe (no idiomatic existing usage relied on the old behavior). Closes [#142](https://github.com/saveenr/VisioAutomation/issues/142).

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
