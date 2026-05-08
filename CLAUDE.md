# CLAUDE.md

Project-specific guidance for Claude Code sessions in this repo. Loaded automatically.

## What this is

A .NET Framework library plus a PowerShell module that automate Microsoft Visio via COM interop. Full picture: [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md).

## The 2026 refresh — read this before making changes

Active branch: `master`. Phase 1 of the refresh has merged; Phase 2 (the final release of the `VisioAutomation2010` NuGet) and Phase 3 (modernization) are still ahead. Work is staged in three phases per [docs/ROADMAP.md](docs/ROADMAP.md):

1. **Phase 1 — VS 2022 cleanup** *(done; merged to master 2026-05-03)*. Stayed on the then-current TFMs (.NET Framework 4.5.2 for shipping libs, .NET Framework 4.7.2 for tests) and shipped the **Visio PowerShell 4.6.1** release as the conclusion of this phase.
2. **Phase 2 — Cut a final release** of the `VisioAutomation2010` NuGet (currently `2.6.0`) with refreshed docs. Two prereqs are deferred and need user discussion: the version-number policy (NuGet `2.6.0` vs PS module `4.6.1`) and a leftover-Visio-process flakiness investigation.
3. **Phase 3 — Modernization.** Move to VS 2026 (which requires bumping TFMs to 4.7.2), modern C#, possibly modern .NET, automated releases.

When in doubt whether a change fits the current phase, check [docs/ROADMAP.md](docs/ROADMAP.md).

## Build prerequisites

The shipping libs target .NET Framework 4.5.2. The reference assemblies are supplied by the [`Microsoft.NETFramework.ReferenceAssemblies.net452`](VisioAutomation_2010/Directory.Packages.props) NuGet package — **no .NET Framework Developer Pack install required** (this changed during Phase 3's SDK migration; before that, the 4.5.2 Developer Pack had to be installed via chocolatey).

If you see `MSB3644: The reference assemblies for .NETFramework,Version=v4.5.2 were not found`, NuGet restore didn't run; run `msbuild -t:Restore` first.

## Build commands (verified working)

```bash
MSBUILD="/c/Program Files/Microsoft Visual Studio/2022/Community/MSBuild/Current/Bin/MSBuild.exe"
"$MSBUILD" VisioAutomation_2010/VisioAutomation2010.sln -t:Restore
"$MSBUILD" VisioAutomation_2010/VisioAutomation2010.sln -p:Configuration=Debug -m
```

- **Do not** use `dotnet build` — projects are still legacy csproj (SDK-style migration is Phase 3 Pass 2).
- **Do not** use VS 2026's MSBuild (under `Program Files\Microsoft Visual Studio\18\`) — its .NET Framework floor is 4.6.2 and the shipping libs are on 4.5.2. This is a Phase 3 unblocker (deferred until after the LTSB 2016 sunset on 2026-10-13).

Package versions live in [`VisioAutomation_2010/Directory.Packages.props`](VisioAutomation_2010/Directory.Packages.props) (Central Package Management); individual csprojs use versionless `<PackageReference>` items.

Full reference (IDE flow, test invocation): [docs/BUILDING.md](docs/BUILDING.md).

## Tests need a live Visio

All test projects exercise real Visio COM calls. There is no mock/fake layer (intentional). Tests cannot be run on a machine without Microsoft Visio installed — flag this rather than claiming a green run.

## Per-commit conventions

- **Changelogs:** when a change is consumer-visible (public API, behavior, supported runtime, dependencies) add an entry to the matching `[Unreleased]` section of [`NuGet/CHANGELOG.md`](NuGet/CHANGELOG.md) or [`VisioAutomation_2010/VisioPowerShell/CHANGELOG.md`](VisioAutomation_2010/VisioPowerShell/CHANGELOG.md) in the **same commit**. Pure internal / build / docs changes don't need entries.
- **PowerShell loader scripts** in `VisioAutomation_2010/VisioPowerShell/`: `Load*` does in-session imports; `Install*` does persistent installs to the user's PS modules folder; `*.ISE.ps1` is the ISE-launched variant. See the folder's [README.md](VisioAutomation_2010/VisioPowerShell/README.md).

## Tooling notes

- **Shell:** Windows host. Both Bash and PowerShell are available. Use Bash for git and Unix-style tooling; use PowerShell for `.ps1` parse checks (`[System.Management.Automation.PSParser]::Tokenize`) and Windows-specific operations.
- **GitHub access:** the `TheSevenPens` git identity has push access to all three repos in play (this repo, [`VisioAutomation_GitBook_Docs`](https://github.com/saveenr/VisioAutomation_GitBook_Docs), and [`VisioPowerShellDocs`](https://github.com/saveenr/VisioPowerShellDocs)). The user-facing docs live in those last two repos as siblings of this repo (cloned to `C:\Users\savee\Documents\GitHub\VisioAutomation_GitBook_Docs\` and `C:\Users\savee\Documents\GitHub\VisioPowerShellDocs\`). PS docs use a version-pinned branch (`visiops_v4_docs`), not master — see [reference_doc_repos.md memory](../../.claude/projects/C--Users-savee-Documents-GitHub-VisioAutomation/memory/reference_doc_repos.md).

## Current state (resume here)

**Phase 1 done. Visio PowerShell module 4.6.1 shipped to PSGallery on 2026-05-03** (tag `VisioPS_4.6.1`). The `2026_Refresh` feature branch has been fast-forwarded into `master` and deleted; new work goes on `master` directly or on a fresh feature branch.

The 4.6.1 release bundles all the Phase 1 cleanup work plus four cmdlet bug fixes (`Lock-VisioShape` / `Unlock-VisioShape` switches now actually bind, `Export-VisioShape` no longer trips on its inverted file-existence check, `New-VisioShape` polyline / Bezier minimum-point validation actually throws). NuGet was deliberately not bumped — all four fixes are in `VisioPowerShell/Commands/`, not in the underlying library — so NuGet stays at `2.6.0` and the version-divergence policy decision is still deferred to Phase 2.

The first publish run surfaced several PSGallery / PS 5.1 gotchas (TLS 1.2 default, PowerShellGet 1.x silent-error bug, PS 5.1 vs 7 user-module path divergence, .ps1 file encoding). All are documented in [VisioPowerShellDocs/developer-info/publishing-to-powershell-gallery.md](https://saveenr.gitbook.io/visiopowershell/developer-info/publishing-to-powershell-gallery) and worked around by [`Publish-VisioPSToGallery.ps1`](VisioAutomation_2010/VisioPowerShell/Publish-VisioPSToGallery.ps1).

**Doc audit:** complete on both gitbook repos. Reader-facing summaries at [VisioPowerShellDocs/documentation-changes.md](https://saveenr.gitbook.io/visiopowershell/documentation-changes) and [VisioAutomation_GitBook_Docs/documentation-changes.md](https://saveenr.gitbook.io/visioautomation/documentation-changes). Every cmdlet has a complete page in the standard `## Syntax` / `## Parameters` / `## Examples` / `## See also` layout, and the .NET-side helper-class coverage was extended in three tiers (Tier 1: Hyperlinks / Lock cells / Control handles / Connection points / Connectors. Tier 2: Shape format-layout-xform / Page cells / Text formatting / Geometry / Application. Tier 4: Analyzers / Visio error log / UndoScope / Exceptions / full Extension-methods rewrite). Only Tier 3 (the `VisioAutomation.Models` project) is still pending and is recorded in [`docs/futures/docs.md`](docs/futures/docs.md#expand-net-side-doc-coverage--tier-3-visioautomationmodels) under *"Expand .NET-side doc coverage — Tier 3"*.

**Repo state:** all three repos in sync with origin, working trees clean. No in-flight branches.

**Test infrastructure:** all 177 tests across the four test projects pass as of 2026-05-04 (VTest 94, VTest.Models 45, VTest.Scripting 34, VTest.PowerShell 4). Each test run leaves zero Visio orphan processes. Fixes that landed to get there:

- `12027821` &mdash; removed the legacy MSTest v1 project-type GUID (`{3AC096D0-…}`) from all four test csprojs (was making VS Test Explorer use the legacy discovery path). Added the `System.Threading.Tasks.Extensions 4.5.4` package as a transitive dep of MSTest's runner.
- `5606adcc` &mdash; enabled `<AutoGenerateBindingRedirects>` + `<GenerateBindingRedirectsOutputType>` so library projects emit `.dll.config` files with binding redirects.
- `5cbf11cd` &mdash; removed redundant `[DeploymentItem]` attributes (8 of them across `XmlErrorLogTests`, `DrawModel_DirectedGraph`, `DrawModel_OrgChartTests`). The data files are already `CopyToOutputDirectory=Always`, so the attributes were redundant; their only effect was triggering VS Test Explorer's deployment-mode behavior, which dropped runtime dependencies on the floor.
- `fb1799d4` &mdash; reverted an attempt to use a solution-level `default.runsettings` with `DeploymentEnabled=false` (it broke previously-passing tests in VS Test Explorer; not the right approach).
- `da9bba0a` &mdash; fixed `Dom_DrawOrgChart` by version-guarding the template filename (`orgchart.vst` for Visio &lt; v15, `orgch_u.vstx` for v15+). Visio 2013 replaced binary `.vst` templates with XML-based `.vstx` and modern Visio installs only ship `.vstx`.
- `b77a99f0` &mdash; **enabled 14 silently-skipped tests** by adding `[MUT.TestClass]` to seven test classes that derived from `Framework.VTest` but lacked the attribute (MSTest 4.x doesn't inherit `[TestClass]` from a base class). The build emitted no warning, so the regression was invisible. Test count went 163 &rarr; 177. Same commit fixed the `OrgChartStyling.cs:9` production bug surfaced by the now-running tests (`Visio2013Template = "orgch_u.vst"` &rarr; `"orgch_u.vstx"`).
- `9a592a9d` &mdash; fixed Visio-process orphan leak. Each testhost was leaking its `Framework.VTest.app_ref` singleton on exit (4 orphans per clean run, ~945 MB; 18 orphans / 4.5 GB after re-runs). Added `[AssemblyCleanup]` hooks per project that close all docs forcibly then `app.Quit(true)` (mirrors the production `ApplicationCommands.cs` pattern). Refactored 3 rogue tests in `DrawModel_OrgChartTests.cs` to use the singleton instead of spawning a second Visio.

## Resume here

**Planning is now semester-based.** [`docs/MILESTONES.md`](docs/MILESTONES.md) is the canonical forward-looking work plan, organized by quarter (CY26Q2, CY26Q3, CY26Q4, CY27Q1, CY27Q2 &mdash; with matching [GitHub milestones](https://github.com/saveenr/VisioAutomation/milestones)). Themed milestones (A&ndash;H) tag each item by *what kind of work* it is; the semester is *when*. Twelve meta-task issues filed (#149&ndash;#160). Existing docs issues (#131&ndash;#133) and identity issues (#146&ndash;#148) re-tagged to their semesters. **Guiding principle:** improve before audience-reducing changes &mdash; 2026 is dedicated to docs / ergonomics / tests / identity; 2027+ is when audience-reducing modernization (TFM bumps, modern .NET, Visio 2013 baseline) begins.

**Identity transition (Saveen / saveenr &rarr; SevenPens) largely done.** 7 of 9 axes complete: commit author identity, nuget.org publishing, PSGallery publishing, displayed authorship in metadata, code-comment refs, LICENSE.txt brand swap, test-fixture scrub. Axis 5 (hosting URLs) is scheduled to CY26Q4 (issues [#146](https://github.com/saveenr/VisioAutomation/issues/146) GitHub repo move + [#147](https://github.com/saveenr/VisioAutomation/issues/147) gitbook moves + Phase 5b in-repo URL rewrite). Axis 9 ([#148](https://github.com/saveenr/VisioAutomation/issues/148) retire unused VisioAutomation legacy account) ride-along on CY26Q4. Detail: [`docs/futures/identity.md`](docs/futures/identity.md).

**Releases shipped this session:** **VisioAutomation NuGet 3.0.0** to nuget.org ([package page](https://www.nuget.org/packages/VisioAutomation2010/3.0.0)) &mdash; first end-to-end CI publish via the freshly-shipped `publish-nuget.yml`. Surfaced and worked around a Microsoft-package compliance gate on the saveenr account; SevenPens is now the canonical publisher for both NuGet and PSGallery. **Visio PowerShell 4.7.2** to PSGallery (already shipped; current). The CI release flow (`release-{psmodule,nuget}.yml` + `publish-{psmodule,nuget}.yml`) is fully operational. Detail in [`docs/futures/releases.md`](docs/futures/releases.md).

**Repo state:** master at HEAD (`98753108`). Working trees clean across all three repos (this repo + the two gitbook repos). PSGallery: `Visio` 4.7.2 live. NuGet: `VisioAutomation2010` 3.0.0 live. [`NuGet/CHANGELOG.md`](NuGet/CHANGELOG.md) `[Unreleased]` has the SevenPens brand-swap entry plus the [#82](https://github.com/saveenr/VisioAutomation/issues/82) `MsaglRenderer.format_shape` fix (next NuGet release pickup); [`VisioPowerShell/CHANGELOG.md`](VisioAutomation_2010/VisioPowerShell/CHANGELOG.md) `[Unreleased]` is empty.

**Test infrastructure:** all four projects green &mdash; VTest 100/100, VTest.Models 64/64, VTest.Scripting 44/44, VTest.PowerShell 22/22. Total **230 tests** at runtime (226 enumerated by single-line `[TestMethod]` regex; +4 are multi-line attributes that the regex misses). Each test run still leaves zero Visio orphan processes. Naming convention now consistent across the suite (full rollout in this session per [#165](https://github.com/saveenr/VisioAutomation/issues/165)&ndash;[#169](https://github.com/saveenr/VisioAutomation/issues/169)); forward-looking convention codified in [`docs/TESTING.md`](docs/TESTING.md#naming-conventions).

**Last session (2026-05-07) summary:** 7 issues closed ([#80](https://github.com/saveenr/VisioAutomation/issues/80), [#82](https://github.com/saveenr/VisioAutomation/issues/82), [#102](https://github.com/saveenr/VisioAutomation/issues/102), [#105](https://github.com/saveenr/VisioAutomation/issues/105), [#149](https://github.com/saveenr/VisioAutomation/issues/149), [#150](https://github.com/saveenr/VisioAutomation/issues/150), [#153](https://github.com/saveenr/VisioAutomation/issues/153)), 3 new filed ([#161](https://github.com/saveenr/VisioAutomation/issues/161) version-compat tables, [#162](https://github.com/saveenr/VisioAutomation/issues/162) Phase A scoping, [#163](https://github.com/saveenr/VisioAutomation/issues/163) Phase B implementation), 6 commits to master, 1 follow-up still pending ([#117](https://github.com/saveenr/VisioAutomation/issues/117) revisit ~2026-05-27 then [#151](https://github.com/saveenr/VisioAutomation/issues/151) closes). New ADR pattern: [`docs/decisions/`](docs/decisions/) folder established with [`tests-need-visio.md`](docs/decisions/tests-need-visio.md) as the inaugural entry. PSVA gap audit landed in [`docs/futures/build-and-code.md`](docs/futures/build-and-code.md).

**Previous session (2026-05-07b) summary:** Big session, four threads of work.

1. **[#163](https://github.com/saveenr/VisioAutomation/issues/163) attempted, abandoned, deferred CY26Q2 &rarr; CY26Q3.** Pipeline parameter set on `Connect-VisioShape`. Cmdlet itself works (verified manually in PS 5.1), but regression tests routed through `VTest.PowerShell`'s `InvokeScript<T>` runspace plumbing failed with `CommandTarget: application does not match doc.application`, distinct from cmdlet logic. Surfaced the underlying constraint: `VisioCmdlet : SMA.Cmdlet` (not `PSCmdlet`) was a deliberate early-on choice for testability reasons, and several cmdlet designs (incl. detecting `ParameterSetName`) want PSCmdlet's surface. Rather than work around it for #163 alone, filed [#164](https://github.com/saveenr/VisioAutomation/issues/164) "Investigate switching VisioCmdlet base from Cmdlet to PSCmdlet" (CY26Q3, Milestone B) and parked #163 behind it. All in-progress code reverted; the constraint is captured in the [feedback_pscmdlet_avoid memory](../../.claude/projects/C--Users-savee-Documents-GitHub-VisioAutomation/memory/feedback_pscmdlet_avoid.md).

2. **[#161](https://github.com/saveenr/VisioAutomation/issues/161) closed.** Two new gitbook pages live: [VA NuGet version compatibility](https://saveenr.gitbook.io/visioautomation/version-compatibility) and [VisioPS module version compatibility](https://saveenr.gitbook.io/visiopowershell/developer-info/version-compatibility). Each row's data sourced from per-tag `csproj` / `nuspec` / `Visio.psd1`. Linked from [`readme.md`](readme.md)'s Documentation section and from both `CHANGELOG.md` preambles. PS-side page also documents that `Visio.psd1`'s historical `PowerShellVersion = '2.0'` claim is stale: actual minimum is PS 5.1 on net452.

3. **[#152](https://github.com/saveenr/VisioAutomation/issues/152) Phase H1 partial.** New [`docs/RELATED-REPOS.md`](docs/RELATED-REPOS.md) drafted with status calls (preliminary) for the 9 sibling Visio repos under `saveenr`. Two rows confirmed (`VisioAutomation2007` self-archive, `visio-templates` empty); 7 others tagged **(preliminary)** awaiting owner confirmation. Per-repo README touches still pending. Kept #152 open with a [comment](https://github.com/saveenr/VisioAutomation/issues/152#issuecomment-4402752238) listing the 7 status calls needed.

4. **[#132](https://github.com/saveenr/VisioAutomation/issues/132) closed.** Tier 3 of the .NET-side doc audit landed: 6 new pages under `models/` on `VisioAutomation_GitBook_Docs` covering the entire `VisioAutomation.Models` project ([Declarative DOM](https://saveenr.gitbook.io/visioautomation/models/dom), [Layouts](https://saveenr.gitbook.io/visioautomation/models/layouts), [Directed graph](https://saveenr.gitbook.io/visioautomation/models/directed-graph), [Layout styles](https://saveenr.gitbook.io/visioautomation/models/layout-styles), [Org charts](https://saveenr.gitbook.io/visioautomation/models/org-charts), [Form pages](https://saveenr.gitbook.io/visioautomation/models/forms)). Each page leads with a working snippet adapted from `VTest.Models/`. **The .NET-side gitbook is now fully Tier 1+2+3+4 covered.**

5. **CY26Q3 review &rarr; CY26Q2 promotions.** Triaged the 10 CY26Q3 items, moved [#132](https://github.com/saveenr/VisioAutomation/issues/132) (closed in this session) and [#154](https://github.com/saveenr/VisioAutomation/issues/154) (test-coverage audit) up to CY26Q2. The other 8 stayed in Q3 (most are blocked or coupled to a user decision).

6 commits to master across this session (`bea0ee89` -> `c0869eea`); 9 commits across the two gitbook repos.

**Dev-environment note:** Smart App Control on Windows 11 can transiently block freshly-built test DLLs (HRESULT 0x800711C7). The block clears in 30&ndash;60s as SAC's cloud reputation cache catches up. Durable fix is to turn SAC off (one-way operation; can't re-enable without reinstalling Windows). CI runners are not subject to this.

**This session (2026-05-08) summary:** Single-theme session, focused on test naming and convention rollout. Six issues closed in sequence:

1. **[#154](https://github.com/saveenr/VisioAutomation/issues/154)** &mdash; test-coverage gap audit. Output: [`docs/futures/test-coverage-gaps.md`](docs/futures/test-coverage-gaps.md), a prioritized gap list across the public API surface (~340 types) cross-referenced against the test inventory (211 enumerated). Headline finding: PowerShell cmdlet coverage is the biggest gap (~56 of 70 cmdlets without dedicated tests, exactly the regression class that shipped in module 4.6.1).
2. **[#165](https://github.com/saveenr/VisioAutomation/issues/165)** &mdash; worst-offender renames + forward-looking naming convention codified in [`docs/TESTING.md`](docs/TESTING.md#naming-conventions). 36 method renames + 3 class renames + 3 file renames covering numbered-suffix series, redundant `Test_` infix, mis-labeled singletons.
3. **[#166](https://github.com/saveenr/VisioAutomation/issues/166)** &mdash; class-redundant prefix sweep + file-name underscore cleanup across VTest.Scripting and VTest.Models. ~80 method prefix drops + 30 file/class renames.
4. **[#167](https://github.com/saveenr/VisioAutomation/issues/167)** &mdash; VTest.PowerShell sweep, vague-name improvements, stylistic outliers, helper-method visibility tightening. ~21 helpers made private; 15 vague names improved with body-inspected `Subject_Scenario_ExpectedOutcome` names.
5. **[#168](https://github.com/saveenr/VisioAutomation/issues/168)** &mdash; "Scenarios" kitchen-sink splits. Four kitchen-sink tests split into 13 (Hyperlinks, Controls, Selection, CustomProps); three mis-labeled singletons just renamed (Undo, CloseDocument, ConnectionPoints).
6. **[#169](https://github.com/saveenr/VisioAutomation/issues/169)** &mdash; final pass: VTest project prefix sweep + AddRemove paired-action splits. 22 prefix drops + 3 AddRemove tests split into 9.

After all the rename work, manually ran the full suite for the first time today; surfaced two pre-existing `doc.Close()` (no `force=true`) bugs in [`DirectedGraphDrawModelTests.cs`](VisioAutomation_2010/VTest.Models/DirectedGraphDrawModelTests.cs) (Drawing9 and Drawing20 save prompts that needed manual `Don't Save` clicks). Fixed in `98753108` by switching the 5 occurrences to `doc.Close(true)` and adding `using VisioAutomation.Extensions;`.

Test count: 211 enumerated &rarr; 226 enumerated (230 runtime). 8 commits to master across this session (`5946629a` &rarr; `98753108`).

Plus filed [#170](https://github.com/saveenr/VisioAutomation/issues/170) for next session focus (LINQ provider for ShapeSheets).

## Next session priorities (CY26Q2)

**Start next session with [#170](https://github.com/saveenr/VisioAutomation/issues/170) &mdash; LINQ provider for ShapeSheet queries (experimental spike).** Filed at the end of this session at user request.

**Goal**: working experimental LINQ-shaped API for ShapeSheet queries on a feature branch, demonstrating ergonomics vs the existing `CellQuery` / `SectionQuery` builder API. **Not** a commitment to ship. Output is a writeup with a recommend / don't-recommend / pursue-with-modifications conclusion. Read the issue body for the full scope and acceptance criteria.

**Background to load before starting:**

- Existing query API:
  - [`VisioAutomation/ShapeSheet/Query/CellQuery.cs`](VisioAutomation_2010/VisioAutomation/ShapeSheet/Query/CellQuery.cs) (140 lines) &mdash; build a list of `SrcConstants` columns, call `GetFormulas(page, shapeids)` or `GetResults<T>(...)`. Returns `string[][]` / `T[][]` keyed by shape and column.
  - [`VisioAutomation/ShapeSheet/Query/SectionQuery.cs`](VisioAutomation_2010/VisioAutomation/ShapeSheet/Query/SectionQuery.cs) (292 lines) &mdash; same idea for variable-row sections (custom properties, hyperlinks, geometry rows).
- Existing tests exercising the API (good targets for the side-by-side comparison):
  - [`VTest/Core/ShapeSheet/ShapeSheetQueryTests.cs`](VisioAutomation_2010/VTest/Core/ShapeSheet/ShapeSheetQueryTests.cs) &mdash; 7 tests covering single/multi-shape, section-row handling, out-of-order shape ids.
  - [`VTest/Core/ShapeSheet/ShapeSheetWriterTests.cs`](VisioAutomation_2010/VTest/Core/ShapeSheet/ShapeSheetWriterTests.cs) &mdash; 6 tests on the symmetric writer side (out of scope for this spike but useful for context).
- The 17 cell-record types in [`VisioAutomation/Shapes/`](VisioAutomation_2010/VisioAutomation/Shapes/), [`VisioAutomation/Pages/`](VisioAutomation_2010/VisioAutomation/Pages/), [`VisioAutomation/Text/`](VisioAutomation_2010/VisioAutomation/Text/) (`*Cells.cs` deriving from `CellRecord`). These are typed-property views over groups of cells; they may or may not be the right base for a LINQ-shaped API.

**Suggested approach (the issue body has more):**

1. Create a feature branch named `experiment/linq-shapesheet`.
2. Start with the **lightweight** option first &mdash; `IEnumerable<T>` extensions over `Page.Shapes` exposing typed cell properties, with LINQ-to-objects projecting to formulas/results, internally calling the existing `CellQuery`. Faster to spike than a full `IQueryable<T>` provider, and may already settle the ergonomic question.
3. Sandbox the spike code in a new project (e.g. `VPlayground.Linq` exe under `VisioAutomation_2010/`) so the shipping `VisioAutomation` library stays clean during the spike.
4. Reimplement at least 3 representative queries from `ShapeSheetQueryTests.cs` side-by-side with the original.
5. Write the comparison up at `docs/futures/linq-shapesheet-spike.md` (or as a comment on [#170](https://github.com/saveenr/VisioAutomation/issues/170) if shorter).
6. Close [#170](https://github.com/saveenr/VisioAutomation/issues/170) with the recommendation, regardless of which way it points.

**Lower priority for the same semester** (unchanged from previous session):

- **[#152](https://github.com/saveenr/VisioAutomation/issues/152) Visio repo portfolio audit Phase H1**: centralized index landed earlier as [`docs/RELATED-REPOS.md`](docs/RELATED-REPOS.md). The other half of acceptance (per-repo `README.md` updates on each of the 9 sibling repos) is blocked on owner-confirmed status calls for 7 rows still tagged **(preliminary)**. See the [comment thread on #152](https://github.com/saveenr/VisioAutomation/issues/152#issuecomment-4402752238) for the open status questions. Once those are confirmed, the per-repo README pass is mechanical: 9 small commits on 9 sibling repos.
- **[#151](https://github.com/saveenr/VisioAutomation/issues/151) triage tracker**: open intentionally as the home for the [#117](https://github.com/saveenr/VisioAutomation/issues/117) revisit. **Calendar reminder: ~2026-05-27** &mdash; if `@tcox8` hasn't confirmed [Visio PS 4.7.0](https://github.com/saveenr/VisioAutomation/releases/tag/VisioPS_4.7.0) fixed their case by then, close [#117](https://github.com/saveenr/VisioAutomation/issues/117) as fixed-by-#144, then close [#151](https://github.com/saveenr/VisioAutomation/issues/151).

**Out of scope unless we finish CY26Q2 with bandwidth to spare:** anything in [CY26Q3](https://github.com/saveenr/VisioAutomation/milestone/3). Eight items there: [#131](https://github.com/saveenr/VisioAutomation/issues/131) (`VisioScripting.Client` docs &mdash; **blocked on [#156](https://github.com/saveenr/VisioAutomation/issues/156)**), [#133](https://github.com/saveenr/VisioAutomation/issues/133) (troubleshooting page, deliberately reactive), [#155](https://github.com/saveenr/VisioAutomation/issues/155) (doc-sample-as-test discussion, couples to [#157](https://github.com/saveenr/VisioAutomation/issues/157)), [#156](https://github.com/saveenr/VisioAutomation/issues/156) (decide `VisioScripting` public-API status &mdash; making this call would unblock [#131](https://github.com/saveenr/VisioAutomation/issues/131)), [#157](https://github.com/saveenr/VisioAutomation/issues/157) (decide long-term docs location), [#162](https://github.com/saveenr/VisioAutomation/issues/162) (PSVA Phase A scoping &mdash; couples to [#164](https://github.com/saveenr/VisioAutomation/issues/164)), [#163](https://github.com/saveenr/VisioAutomation/issues/163) (parked behind [#164](https://github.com/saveenr/VisioAutomation/issues/164)), [#164](https://github.com/saveenr/VisioAutomation/issues/164) (PSCmdlet investigation, freshly filed).

**Possible side quests** (small, opportunistic, not on any milestone but easy to pick up):
- Trim the redundant `<frameworkAssembly assemblyName="Microsoft.Office.Interop.Visio" />` line in [`NuGet/VisioAutomation2010.nuspec`](NuGet/VisioAutomation2010.nuspec) &mdash; redundant with the bundled-PIA-via-NuGet behavior since the SDK migration.
- Open a NuGet support case to declassify the `saveenr` account (low priority; the SevenPens workaround works fine, but symmetry with the gitbook / GitHub repo identity would be nicer once Axis 5 lands).

## Other docs in this repo

- [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md) — projects, dependencies, central concepts
- [docs/BUILDING.md](docs/BUILDING.md) — full build / test / install reference (incl. dev-pack install commands)
- [docs/TESTING.md](docs/TESTING.md) — test-suite design and conventions (shared `Framework.VTest` base, per-testhost Visio singleton, `[AssemblyCleanup]` orphan-prevention, MSTEST0030 enforcement). Per-project READMEs sit next to each test csproj.
- [docs/GLOSSARY.md](docs/GLOSSARY.md) — Visio and codebase terminology
- [docs/ROADMAP.md](docs/ROADMAP.md) — staged plan (Phase 1 / 2 / 3); read first for orientation
- [docs/MILESTONES.md](docs/MILESTONES.md) — themed milestones with proposed target windows. Sits between ROADMAP.md (phases) and FUTURES.md (backlog index)
- [docs/FUTURES.md](docs/FUTURES.md) — index of backlog items, split by topic into [`docs/futures/*.md`](docs/futures/)
- [docs/decisions/](docs/decisions/) — architectural decision records (ADRs); one file per long-lived structural choice. Inaugural entry: [`tests-need-visio.md`](docs/decisions/tests-need-visio.md)
- [docs/OVERVIEW.md](docs/OVERVIEW.md) — entry point, links to the above
