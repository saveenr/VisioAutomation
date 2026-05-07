# Changelog — VisioAutomation NuGet

All notable changes to the [`VisioAutomation2010`](https://www.nuget.org/packages/VisioAutomation2010/) NuGet package are documented here.

This package bundles `VisioAutomation.dll`, `VisioAutomation.Models.dll`, `VisioScripting.dll`, plus the supporting `Microsoft.Msagl.dll` and `GenTreeOps.dll`. For the related PowerShell module's release history see [`VisioAutomation_2010/VisioPowerShell/CHANGELOG.md`](../VisioAutomation_2010/VisioPowerShell/CHANGELOG.md).

The format follows [Keep a Changelog 1.1.0](https://keepachangelog.com/en/1.1.0/) and the package follows [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

> **Contributors:** when a change affects what consumers of this NuGet package see (public API, behavior, supported runtime), add an entry under `[Unreleased]`. Pure internal/build/docs changes don't need to appear here.

## [Unreleased]

### Fixed
- `MsaglRenderer` now honors `Node.Size` when `Node.Cells` is also set on a directed-graph node. Previously the `Cells` assignment overwrote the entire `shape_node.Cells` object, silently dropping the `XFormWidth` / `XFormHeight` populated from `Size`, so the rendered shape came out at the master's default size instead of the requested `Size`. Now the user's `Cells` is merged onto the existing cells via `ApplyFormulasTo`, which preserves Size-derived width/height. Closes [#82](https://github.com/saveenr/VisioAutomation/issues/82).

  Precedence is now field-by-field rather than all-or-nothing. Only the styling-only row's behavior changes; all other scenarios are unaffected:

  | Scenario | Pre-fix | Post-fix |
  |---|---|---|
  | Only `Size` set | `Size` | `Size` |
  | Only `Cells.XFormWidth/Height` set | `Cells` | `Cells` |
  | Both `Size` AND `Cells.XFormWidth/Height` set | `Cells` | `Cells` |
  | `Size` set + `Cells` styling-only (no width/height) | **Master default** (bug) | **`Size`** (fix) |
  | Neither set | Master default | Master default |

### Changed
- Package metadata's `<authors>` and `<copyright>` fields updated from `saveenr` / `Copyright Saveen Reddy` to `SevenPens` / `Copyright SevenPens` to reflect the new dev-team identity. No functional change; the displayed-author string on the [nuget.org package page](https://www.nuget.org/packages/VisioAutomation2010) updates on the next release. Legal copyright record (LICENSE.txt) updated correspondingly.

## [3.0.0] - 2026-05-07

Major version bump. Several behavior changes that affect callers:

- The exception type thrown by `CustomPropertyHelper.Set` and `UserDefinedCellHelper.Set` for malformed values changes from `COMException` to `ArgumentException` (with the original `COMException` preserved as `InnerException`).
- `Client.Model.DrawDirectedGraphDocument` now respects each layout's `UseDynamicConnectors` setting from XML, instead of hardcoding it to `false`.
- `DirectedGraphDocumentLoader.LoadFromXml` now rejects non-`<directedgraph>` root elements that previously parsed silently.
- Minimum supported .NET Framework rises from 4.0 to 4.5.2 (binary-breaking for consumers still on net40).

Source compatibility for property accesses is preserved: the new `Formula` property on `CustomPropertyCells` and `UserDefinedCellCells` is paired with an `[Obsolete]` `Value` alias scheduled for removal in a later release.

### Changed
- Minimum required .NET Framework raised from 4.0 to **4.5.2**. .NET Framework 4.5.2 was released in 2014 and has shipped via Windows Update for years, so consumers on a current Windows machine are essentially unaffected. The package's `lib\` folder also moves from `lib\net40` to `lib\net452`, correctly reflecting the binaries' actual target framework — the previous `lib\net40` claim would silently install the assemblies into a `net40` project where they would fail to load.
- Replaced the deprecated `<licenseUrl>` nuspec element with the modern SPDX `<license type="expression">MIT</license>` form. The license itself is unchanged (still MIT) — only the metadata representation differs, which is what nuget.org's package page now expects.
- The package now includes a `README.md` at its root (sourced from the repo's [`readme.md`](https://github.com/saveenr/VisioAutomation/blob/master/readme.md)) so the package page on nuget.org renders project content directly instead of just metadata.

### Added
- `DirectedGraphDocumentLoader.LoadFromXml` now honors a `connectortype` attribute on `<renderoptions>`, accepting `Curved` (default), `Straight`, or `RightAngle`. Previously the connector type was hardcoded to `Curved` regardless of XML, so every directedgraph rendered from XML had curved connectors. Closes [#140](https://github.com/saveenr/VisioAutomation/issues/140) (sub-issue of [#105](https://github.com/saveenr/VisioAutomation/issues/105)).
- `DirectedGraphDocumentLoader.LoadFromXml` now honors a `direction` attribute on `<renderoptions>`, accepting `TopToBottom` (default), `BottomToTop`, `LeftToRight`, or `RightToLeft`. Also accepts a `layout` attribute, currently `Sugiyama`-only (any other value raises `ArgumentException` with a descriptive message). Closes [#141](https://github.com/saveenr/VisioAutomation/issues/141) (sub-issue of [#105](https://github.com/saveenr/VisioAutomation/issues/105)).
- `DirectedGraphLayout.LayoutOptions` (new public `MsaglOptions` property) carries per-page layout settings parsed from `<renderoptions>`. Defaults to a fresh `MsaglOptions()` for programmatically constructed layouts.
- Typed setters on `CustomPropertyCells`: `SetString(string)`, `SetNumber(int)`, `SetNumber(double)`, `SetBool(bool)`, `SetDate(DateTime)`, `SetFormula(string)`. Each emits a correctly-encoded Visio formula and sets `Type` to match. New canonical replacement for raw `Formula = ...` assignment, which is a foot-gun for string values (a bare identifier like `"testVal"` is parsed as a name reference, not a literal). Closes [#144](https://github.com/saveenr/VisioAutomation/issues/144).
- Typed setters on `UserDefinedCellCells`: `SetString(string)`, `SetFormula(string)`. Same idea, simpler surface (no `Type` field). Closes [#144](https://github.com/saveenr/VisioAutomation/issues/144).
- `Formula` property on `CustomPropertyCells` and `UserDefinedCellCells` (renamed from `Value` to surface that the cell stores a Visio formula, not a literal value).

### Changed
- `CustomPropertyHelper.Set` and `UserDefinedCellHelper.Set` now wrap Visio's formula-error `COMException` (`#NAME?`, `#VALUE!`, etc.) in an `ArgumentException` with a self-explanatory message pointing at `SetString` / `EncodeValues`. **Behavior change:** callers that catch `COMException` from these methods will need to catch `ArgumentException` instead (or both). The original `COMException` is preserved as `InnerException`. Part of [#144](https://github.com/saveenr/VisioAutomation/issues/144).

### Deprecated
- `CustomPropertyCells.Value` and `UserDefinedCellCells.Value` are now `[Obsolete]` aliases for `Formula`. Source compatibility preserved through the deprecation window; the alias is scheduled for removal as part of the Phase 3 modernization work. Migration: rename `cells.Value` to `cells.Formula`, or use the new typed setters (`SetString`, etc.). Part of [#144](https://github.com/saveenr/VisioAutomation/issues/144).

### Changed
- `Client.Model.DrawDirectedGraphDocument` now respects each layout's `LayoutOptions` (the per-page `MsaglOptions`) when rendering, instead of always overriding `UseDynamicConnectors` to `false` and ignoring the parsed `usedynamicconnectors` / `scalingfactor` from XML. **Behavior change:** XML `<renderoptions>` attributes that were previously parsed but silently dropped now actually take effect; programmatic callers who relied on the hardcoded `UseDynamicConnectors=false` should set it explicitly on `dg_layout.LayoutOptions` before calling.
- `DirectedGraphDocumentLoader.LoadFromXml` now validates that the root XML element is `<directedgraph>` and throws `ArgumentException` otherwise. This is the same root name the `Import-VisioModel` cmdlet has always required; the bare loader API previously accepted any root and silently parsed only its `<page>` children, which let typos like `<autolayoutdrawing>` succeed without warning. Surfaces the schema mismatch that bit the [#105](https://github.com/saveenr/VisioAutomation/issues/105) reporter. The four in-repo test fixtures (`VTest/datafiles/directed_graph_*.xml`) have been re-rooted from `<autolayoutdrawing>` to `<directedgraph>` to match.

### Fixed
- `OrgChartDocument.Render` no longer fails with `COMException: File not found` on Visio 2013+. The default `OrgChartStyling.Visio2013Template` was hardcoded to `orgch_u.vst`, but Visio 2013 replaced binary `.vst` templates with XML-based `.vstx` and modern Visio installs only ship `orgch_u.vstx`. Updated to `orgch_u.vstx`. Visio 2010 (`Visio2010Template = "orgch_u.vst"`) is unchanged.
- `GeometrySection.Render` now returns the section index Visio assigned to the newly-added geometry section, instead of always returning the literal `0`. This matches the API's stated intent (and the [Geometry](https://saveenr.gitbook.io/visioautomation/geometry) gitbook page's claim) so that callers adding multiple geometry sections can target a specific one by index. Closes [#128](https://github.com/saveenr/VisioAutomation/issues/128).

## Earlier versions

Versions 2.6.0 and earlier predate this changelog. See the [git history](https://github.com/saveenr/VisioAutomation/commits/master/) and [release tags](https://github.com/saveenr/VisioAutomation/releases) for details.
