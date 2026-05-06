# Changelog — VisioAutomation NuGet

All notable changes to the [`VisioAutomation2010`](https://www.nuget.org/packages/VisioAutomation2010/) NuGet package are documented here.

This package bundles `VisioAutomation.dll`, `VisioAutomation.Models.dll`, `VisioScripting.dll`, plus the supporting `Microsoft.Msagl.dll` and `GenTreeOps.dll`. For the related PowerShell module's release history see [`VisioAutomation_2010/VisioPowerShell/CHANGELOG.md`](../VisioAutomation_2010/VisioPowerShell/CHANGELOG.md).

The format follows [Keep a Changelog 1.1.0](https://keepachangelog.com/en/1.1.0/) and the package follows [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

> **Contributors:** when a change affects what consumers of this NuGet package see (public API, behavior, supported runtime), add an entry under `[Unreleased]`. Pure internal/build/docs changes don't need to appear here.

## [Unreleased]

### Changed
- Minimum required .NET Framework raised from 4.0 to **4.5.2**. .NET Framework 4.5.2 was released in 2014 and has shipped via Windows Update for years, so consumers on a current Windows machine are essentially unaffected. The package's `lib\` folder also moves from `lib\net40` to `lib\net452`, correctly reflecting the binaries' actual target framework — the previous `lib\net40` claim would silently install the assemblies into a `net40` project where they would fail to load.
- Replaced the deprecated `<licenseUrl>` nuspec element with the modern SPDX `<license type="expression">MIT</license>` form. The license itself is unchanged (still MIT) — only the metadata representation differs, which is what nuget.org's package page now expects.
- The package now includes a `README.md` at its root (sourced from the repo's [`readme.md`](https://github.com/saveenr/VisioAutomation/blob/master/readme.md)) so the package page on nuget.org renders project content directly instead of just metadata.

### Added
- `DirectedGraphDocumentLoader.LoadFromXml` now honors a `connectortype` attribute on `<renderoptions>`, accepting `Curved` (default), `Straight`, or `RightAngle`. Previously the connector type was hardcoded to `Curved` regardless of XML, so every directedgraph rendered from XML had curved connectors. Closes [#140](https://github.com/saveenr/VisioAutomation/issues/140) (sub-issue of [#105](https://github.com/saveenr/VisioAutomation/issues/105)).
- `DirectedGraphDocumentLoader.LoadFromXml` now honors a `direction` attribute on `<renderoptions>`, accepting `TopToBottom` (default), `BottomToTop`, `LeftToRight`, or `RightToLeft`. Also accepts a `layout` attribute, currently `Sugiyama`-only (any other value raises `ArgumentException` with a descriptive message). Closes [#141](https://github.com/saveenr/VisioAutomation/issues/141) (sub-issue of [#105](https://github.com/saveenr/VisioAutomation/issues/105)).
- `DirectedGraphLayout.LayoutOptions` (new public `MsaglOptions` property) carries per-page layout settings parsed from `<renderoptions>`. Defaults to a fresh `MsaglOptions()` for programmatically constructed layouts.

### Changed
- `Client.Model.DrawDirectedGraphDocument` now respects each layout's `LayoutOptions` (the per-page `MsaglOptions`) when rendering, instead of always overriding `UseDynamicConnectors` to `false` and ignoring the parsed `usedynamicconnectors` / `scalingfactor` from XML. **Behavior change:** XML `<renderoptions>` attributes that were previously parsed but silently dropped now actually take effect; programmatic callers who relied on the hardcoded `UseDynamicConnectors=false` should set it explicitly on `dg_layout.LayoutOptions` before calling.
- `DirectedGraphDocumentLoader.LoadFromXml` now validates that the root XML element is `<directedgraph>` and throws `ArgumentException` otherwise. This is the same root name the `Import-VisioModel` cmdlet has always required; the bare loader API previously accepted any root and silently parsed only its `<page>` children, which let typos like `<autolayoutdrawing>` succeed without warning. Surfaces the schema mismatch that bit the [#105](https://github.com/saveenr/VisioAutomation/issues/105) reporter. The four in-repo test fixtures (`VTest/datafiles/directed_graph_*.xml`) have been re-rooted from `<autolayoutdrawing>` to `<directedgraph>` to match.

### Fixed
- `OrgChartDocument.Render` no longer fails with `COMException: File not found` on Visio 2013+. The default `OrgChartStyling.Visio2013Template` was hardcoded to `orgch_u.vst`, but Visio 2013 replaced binary `.vst` templates with XML-based `.vstx` and modern Visio installs only ship `orgch_u.vstx`. Updated to `orgch_u.vstx`. Visio 2010 (`Visio2010Template = "orgch_u.vst"`) is unchanged.

## Earlier versions

Versions 2.6.0 and earlier predate this changelog. See the [git history](https://github.com/saveenr/VisioAutomation/commits/master/) and [release tags](https://github.com/saveenr/VisioAutomation/releases) for details.
