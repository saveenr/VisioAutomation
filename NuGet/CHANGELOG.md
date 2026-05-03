# Changelog — VisioAutomation NuGet

All notable changes to the [`VisioAutomation2010`](https://www.nuget.org/packages/VisioAutomation2010/) NuGet package are documented here.

This package bundles `VisioAutomation.dll`, `VisioAutomation.Models.dll`, `VisioScripting.dll`, plus the supporting `Microsoft.Msagl.dll` and `GenTreeOps.dll`. For the related PowerShell module's release history see [`VisioAutomation_2010/VisioPowerShell/CHANGELOG.md`](../VisioAutomation_2010/VisioPowerShell/CHANGELOG.md).

The format follows [Keep a Changelog 1.1.0](https://keepachangelog.com/en/1.1.0/) and the package follows [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

> **Contributors:** when a change affects what consumers of this NuGet package see (public API, behavior, supported runtime), add an entry under `[Unreleased]`. Pure internal/build/docs changes don't need to appear here.

## [Unreleased]

### Changed
- Minimum required .NET Framework raised from 4.0 to **4.5.2**. .NET Framework 4.5.2 was released in 2014 and has shipped via Windows Update for years, so consumers on a current Windows machine are essentially unaffected.

## Earlier versions

Versions 2.6.0 and earlier predate this changelog. See the [git history](https://github.com/saveenr/VisioAutomation/commits/master/) and [release tags](https://github.com/saveenr/VisioAutomation/releases) for details.
