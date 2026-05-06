# Completed — 2026 Refresh Done Work

A historical record of backlog items that have landed. The point is institutional memory: *what was done, why, what was tried, what didn't work*. This file is append-only — entries shouldn't be edited after landing except for typo or link fixes.

Forward-looking work lives in the topic-split backlog under [`futures/`](futures/) (indexed by [`FUTURES.md`](FUTURES.md)). The high-level "what shipped in Phase N" headline summary lives in [`ROADMAP.md`](ROADMAP.md) (under each phase header); this file is the per-item detail behind those bullets.

When something on the backlog completes, the Resolution paragraph moves here verbatim, the body entry is deleted from its `futures/*.md` file, and a one-line bullet is added to the relevant phase summary in `ROADMAP.md`. See [`CONTRIBUTING.md`](../CONTRIBUTING.md) for the convention.

Items are grouped by **phase**, then by the same **category** they had in the backlog (Build & tooling, Code & architecture, Tests, Documentation). Within a category they appear in the order they completed.

---

## Phase 1 — VS 2022 cleanup (done; merged to master 2026-05-03)

Headline summary in [`ROADMAP.md`](ROADMAP.md#phase-1--vs-2022-cleanup-done-merged-to-master-2026-05-03). Full detail per item below.

### Build & tooling

#### Update MSTest off the beta
- **What:** All test projects pinned `MSTest.TestFramework` and `MSTest.TestAdapter` to `2.0.0-beta2`.
- **Resolution:** Upgraded both packages to **4.2.2** (latest stable) across all four test projects. Required bumping VTest from .NET Framework 4.5 → 4.7.2 (MSTest 4.x's floor is 4.6.2); the other test projects were already on 4.7.2. Solution builds cleanly. Test code did not need any changes — MSTest's `[TestMethod]`/`Assert.*` API surface is stable across the version jump. Tests not actually run (need a live Visio).

#### Add CI (build-only)
- **Resolution:** [`.github/workflows/build.yml`](../.github/workflows/build.yml) added. Builds the solution in Debug on every push to `master` / `2026_Refresh` and on every PR. Pinned to VS 2022's MSBuild (`vs-version: '17.0'`) since VS 2026's MSBuild can't resolve the .NET Framework 4.5.2 reference assemblies the shipping libs need. The workflow also installs the .NET Framework 4.5.2 Developer Pack via chocolatey before building (those reference assemblies aren't on the runner image). NuGet packages are cached keyed on the hash of all `packages.config` files. Build status surfaces as a badge in the root README.
- **Tail (still active in the backlog):** Running the tests themselves in CI requires a self-hosted Windows runner with Visio installed; tracked under *Run tests in CI* in [`futures/build-and-code.md`](futures/build-and-code.md#run-tests-in-ci).

#### Fix the misnamed PowerShell loader script
- **What:** `DownloadFromPowerShellGallery.ps1` did not download from the PowerShell Gallery — it `Import-Module`d the local `bin\Debug` build.
- **Resolution:** Rewrote the script to `Save-Module Visio` from PSGallery into a local `DownloadedModule/` subfolder (gitignored) and `Import-Module` that. Now serves as a one-shot release-verification helper. Renamed the file to [`LoadFromGallery.ps1`](../VisioAutomation_2010/VisioPowerShell/LoadFromGallery.ps1) for parallelism with the other loader scripts; also renamed [`LoadFromBinDebug_ISE.ps1`](../VisioAutomation_2010/VisioPowerShell/LoadFromBinDebug.ISE.ps1) → `LoadFromBinDebug.ISE.ps1` (dot-suffix is more idiomatic than underscore).

### Code & architecture

#### Audit `Internal/` for dead code
- **Resolution:** Audited `VisioAutomation/Internal/` end-to-end. Two clear-cut wins removed: deleted [`TempHelper.cs`](../VisioAutomation_2010/VisioAutomation/Internal/) (orphaned — duplicate of `ShapesheetHelpers` with snake_case names; not even listed in the csproj, so already wasn't being compiled) and removed the dead `[assembly: InternalsVisibleTo("TestVisioAutomation")]` attribute from `AssemblyInfo.cs` (no `TestVisioAutomation` assembly exists; current test projects are `VTest` / `VTest.Scripting`, granted access elsewhere). All other Internal/ types are actively referenced. Build verified clean.

#### Misc cleanups discovered during the Internal/ audit (mostly)
- **Resolution:** Three of the four findings cleaned up:
  - ✅ Moved `[InternalsVisibleTo("VTest")]` and `[InternalsVisibleTo("VTest.Scripting")]` from `Internal/ArraySegmentEnumerator.cs` to `Properties/AssemblyInfo.cs` where they belong.
  - ✅ Deleted `Vtest/FormatStringParserTest.cs` and `VTest/Core/Extensions/AsEnumerableTest.cs` — turned out to be fully orphaned (not in csproj, contained syntax errors and references to long-gone types like `Isotope.Text.FormatStringParser` and a `VisioAutomationTest` base class). Pure dead code.
  - ✅ Removed `VisioAutomation2010.sln.metaproj` from version control and added `*.sln.metaproj` / `*.sln.metaproj.tmp` to `.gitignore` so it gets regenerated on demand instead of accumulating stale paths.
- **Tail (still active in the backlog):** `LinqExtensions` visibility-vs-folder mismatch is tracked under *Move `LinqExtensions` out of `Internal/` (or rename the folder)* in [`futures/build-and-code.md`](futures/build-and-code.md#move-linqextensions-out-of-internal-or-rename-the-folder).

### Tests

#### Investigate flakiness from leftover Visio processes
- **Resolution (`9a592a9d`):** Each testhost was leaking its `Framework.VTest.app_ref` Visio singleton on exit (no `Quit()` ever called). Empirically: 4 orphans / ~945 MB per clean run; 18 orphans / 4.5 GB after a few re-runs. Three rogue tests in `DrawModel_OrgChartTests.cs` compounded by spawning their own `new IVisio.Application()` with cleanup only on the happy path. Fixed by: adding `[AssemblyCleanup]` hooks per test project that close all docs forcibly then `app.Quit(true)` (mirrors the production `ApplicationCommands.cs` pattern, including `AlertResponse=7` to suppress save prompts &mdash; the initial fix using parameterless `Quit()` did hang the testhost on a "Save Drawing18?" dialog before being corrected). Refactored the three rogue tests to use `this.GetVisioApplication()` and close the doc via `app.ActiveDocument.Close(true)` instead of `app.Quit(true)`. Verified: two consecutive full-suite runs leave zero orphan processes; all 177 tests still pass.

### Documentation

#### Add a `CLAUDE.md` at the repo root
- **What:** Project-specific instructions for future Claude Code sessions: build commands, test rules (need Visio installed), where the public API lives, the `2026_Refresh` branch convention.
- **Resolution:** [`CLAUDE.md`](../CLAUDE.md) added. Covers the staged-plan summary, verified build commands, test prerequisites, the per-commit changelog convention, the PS loader-script naming convention, tooling notes (shell choice, MSBuild path), and pointers to the rest of the docs.

#### Add a `CONTRIBUTING.md`
- **Resolution:** [`CONTRIBUTING.md`](../CONTRIBUTING.md) added at the repo root. Covers the active branch convention, setup pointer to BUILDING.md, the live-Visio test rule, code style guidance (don't reformat, no new files unless needed, default to no comments), commit message format, the per-commit changelog convention, and the per-phase scope rules from FUTURES so contributors don't accidentally violate them.

#### Expand the root `readme.md`
- **Resolution:** Rewrote the three-line readme into a proper landing page: NuGet/PSGallery/license badges, the elevator pitch, an artifact table with install commands, two quick-start examples (C# + PowerShell), links to user guides, in-repo developer docs, release notes, contributing, and license.

#### Add a per-project `README.md` for the larger projects
- **Resolution:** READMEs added for all four larger projects: [`VisioAutomation/`](../VisioAutomation_2010/VisioAutomation/README.md), [`VisioAutomation.Models/`](../VisioAutomation_2010/VisioAutomation.Models/README.md), [`VisioScripting/`](../VisioAutomation_2010/VisioScripting/README.md), [`VisioPowerShell/`](../VisioAutomation_2010/VisioPowerShell/README.md). Each covers folder layout, key types where relevant, and pointers to ARCHITECTURE / GLOSSARY / BUILDING / CHANGELOG.

---

## Phase 3 — Modernization *(in progress)*

Headline summary in [`ROADMAP.md`](ROADMAP.md#phase-3--modernization-in-progress). Full detail per item below.

### Build & tooling

#### Migrate from `packages.config` to `PackageReference`
- **What:** Every csproj used the old `packages.config` NuGet model; transitive dependencies didn't flow, package targets needed hardcoded `<Import>` paths, and `dotnet` CLI / SDK-style projects were impossible.
- **Resolution (`86ef3984`, PR [#134](https://github.com/saveenr/VisioAutomation/pull/134)):** All 11 csprojs converted from `packages.config` to versionless `<PackageReference>` items, csprojs still in legacy format. New [`VisioAutomation_2010/Directory.Build.props`](../VisioAutomation_2010/Directory.Build.props) enables Central Package Management; [`Directory.Packages.props`](../VisioAutomation_2010/Directory.Packages.props) centralizes all 10 package versions. PIA preserved bundle-with-NuGet behavior via `<EmbedInteropTypes>false</EmbedInteropTypes>` + `<PrivateAssets>none</PrivateAssets>`. Hardcoded `System.Management.Automation` paths in VTest and VTest.PowerShell (one to the WindowsPowerShell 3.0 reference assemblies, one to the GAC) replaced with the `Microsoft.PowerShell.3.ReferenceAssemblies` package that VisioPowerShell already used. **Dev-pack install requirement gone:** added `Microsoft.NETFramework.ReferenceAssemblies.{net452,net472}` packages — the package's `.targets` overrides `TargetFrameworkRootPath` so MSBuild finds reference assemblies in the package, not on disk. Defers the TFM bump beyond the LTSB 2016 sunset (2026-10-13) without paying the dev-pack pain in the meantime. Eliminated MSTest hardcoded `<Import>` lines, `EnsureNuGetPackageBuildImports` targets, and the explicit `<Analyzer>` ItemGroup. CI workflows: removed the `choco install netfx-4.5.2-devpack` step; switched cache path to `~/.nuget/packages` keyed on `Directory.Packages.props`. Build time on CI dropped from 3-10 min to ~70s.

#### Modernize SDK-style csproj
- **What:** All 11 csprojs were legacy format with explicit `<Compile Include>` lists, multiple PropertyGroups for Debug/Release/Configuration combinations, and hundreds of lines of accumulated cruft (BootstrapperPackage, ClickOnce, Scc, ProductVersion=9.0.30729, OldToolsVersion, FileUpgradeFlags, etc.).
- **Resolution (Pass 2a `9bd06398` PR [#135](https://github.com/saveenr/VisioAutomation/pull/135) + Pass 2b `e053e9ed` PR [#136](https://github.com/saveenr/VisioAutomation/pull/136) + Pass 2c `53e9197c` PR [#137](https://github.com/saveenr/VisioAutomation/pull/137)):** All 11 csprojs converted from legacy to SDK-style (`<Project Sdk="Microsoft.NET.Sdk">`). Net diff across the three sub-passes: **-1,322 lines.** Each csproj dropped from 100-280 lines to 25-50 lines. The `.sln` project type GUIDs updated for all 11 (`{FAE04EC0-...}` → `{9A19103F-...}`); the solution now uniformly references SDK-style projects.
  - **Pass 2a** (4 shipping libraries: VisioAutomation, .Models, VisioScripting, VisioPowerShell). Net: -688 lines. VisioPowerShell preserved `<AssemblyName>VisioPS</AssemblyName>` (required because `Visio.psd1`'s `ModuleToProcess = 'VisioPS.dll'` references that exact name) and explicit `<RootNamespace>VisioPowerShell</RootNamespace>`. Added explicit `<None Include>` for `Visio.psd1` and `Visio.Types.ps1xml` with `CopyToOutputDirectory=PreserveNewest`. **Drive-by deletion:** `VisioPowerShell/Models/NameCellDictionary.cs` was 7-year-old "Work in progress" dead code from PR #96 (2019-03-30) — never in the legacy `<Compile>` list, never compiled, references types that don't exist, no callers. SDK auto-glob surfaced and deleted it.
  - **Pass 2b** (4 test projects: VTest, VTest.Models, VTest.Scripting, VTest.PowerShell). Net: -360 lines. Resolved three deferred test-cleanup angles in one stroke: the long-standing **MSB3270 x86/AnyCPU mismatch warning** (unified to AnyCPU by dropping `<PlatformTarget>x86</PlatformTarget>`); the legacy `<TestProjectType>UnitTest</TestProjectType>` / `<VSToolsPath>` / `<IsCodedUITest>` cruft (deleted by definition); and the **`Vtest.Models.csproj` → `VTest.Models.csproj` filename casing fix** (lowercase 'v' was a long-standing typo working only because Windows is case-insensitive). Preserved `<AutoGenerateBindingRedirects>true</AutoGenerateBindingRedirects>` + `<GenerateBindingRedirectsOutputType>true</GenerateBindingRedirectsOutputType>` on each test csproj (SDK-style auto-enables these for `OutputType=Exe` but not `Library`); confirmed `.dll.config` still emits with redirects. Added `<IsPackable>false</IsPackable>` for safety. VTest's 10 data files use a glob (`<None Include="datafiles\**" />`) instead of 10 explicit entries. MSTest.Analyzers + MSTEST0030 enforcement preserved through the migration (verified via synthetic regression test).
  - **Pass 2c** (3 exes: VSamples, VSamples.Docs, VPlayground). Net: -274 lines. VPlayground preserved the `COMReference` for `stdole`. `<PlatformTarget>x86</PlatformTarget>` dropped from VSamples and VPlayground to match the AnyCPU policy. SDK auto-handles binding redirects for `OutputType=Exe`; `.exe.config` files emit correctly.

### Tests

#### Test-discovery linter (`MSTest.Analyzers` + MSTEST0030 enforcement)
- **What:** The `[TestClass]`-attribute regression that hid 14 silently-skipped tests for years (commit `b77a99f0`) wasn't caught by anything. MSTest 4.x doesn't inherit `[TestClass]` from a base class, and the build emitted no warning. We needed a Roslyn analyzer or build-time check to prevent recurrence — especially before release CI lands, since a pipeline that silently shrinks the test suite is worse than no pipeline.
- **Resolution (`7700acf7`):** Added `MSTest.Analyzers 4.2.2` to all four test projects' `packages.config` (Pass 1 of SDK migration later moved these to `<PackageReference>`) plus an explicit `<Analyzer Include="..\packages\MSTest.Analyzers.4.2.2\analyzers\dotnet\cs\MSTest.Analyzers.dll" />` reference (also auto-handled post-Pass-1). New solution-level [`VisioAutomation_2010/.editorconfig`](../VisioAutomation_2010/.editorconfig) promotes **MSTEST0030** ("Type containing `[TestMethod]` should be marked with `[TestClass]`") from the analyzer's default warning to **error**. Synthetic test confirmed the analyzer fails the build (MSBuild exit 1) when `[TestClass]` is removed; the regression cannot recur silently. Bonus drive-by: the analyzer also caught one MSTEST0017 (assertion args swapped) in `PageHelperTests.cs:101`, fixed in the same commit.

#### Per-project test READMEs and top-level `docs/TESTING.md`
- **What:** None of the four test projects had READMEs explaining what they cover, what fixtures they need, or what state they assume Visio to be in. The shared infrastructure (`Framework.VTest` base class, per-testhost Visio singleton, `[AssemblyCleanup]` orphan-prevention) had no top-level documentation.
- **Resolution (`a41e97bc`):** Added per-project `README.md` for [`VTest/`](../VisioAutomation_2010/VTest/README.md), [`VTest.Models/`](../VisioAutomation_2010/VTest.Models/README.md), [`VTest.Scripting/`](../VisioAutomation_2010/VTest.Scripting/README.md), and [`VTest.PowerShell/`](../VisioAutomation_2010/VTest.PowerShell/README.md). Added top-level [`docs/TESTING.md`](TESTING.md) covering the test-suite design (the three load-bearing constraints: real Visio, sequential execution, per-testhost singleton); shared infrastructure (`Framework.VTest`, `VTestAppRef`, `[AssemblyCleanup]` per-assembly pattern, datafiles convention); MSTest.Analyzers + MSTEST0030 enforcement; how-to-run pointers; and known gotchas. Cross-linked from `CLAUDE.md` and `docs/OVERVIEW.md`.
- **Tail (still active in the backlog):** The *coverage-gaps* angle of the original "General cleanup of the test projects" entry was deliberately not bundled with this work; deferred indefinitely as too open-ended to scope. Tracked under *Test coverage gaps* in [`futures/tests.md`](futures/tests.md#test-coverage-gaps).
