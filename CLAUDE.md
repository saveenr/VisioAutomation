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

**Repo state:** master at HEAD (`b8710d96`). Working trees clean across all three primary repos (this repo + the two gitbook repos); 8 sibling repos each got a one-shot Phase H1 README commit during this session, `visio-templates` was deleted. The `experiment/linq-shapesheet` branch is on origin (1 commit, motivation doc only). PSGallery: `Visio` 4.7.2 live. NuGet: `VisioAutomation2010` 3.0.0 live. [`NuGet/CHANGELOG.md`](NuGet/CHANGELOG.md) `[Unreleased]` has the SevenPens brand-swap entry, the [#82](https://github.com/saveenr/VisioAutomation/issues/82) `MsaglRenderer.format_shape` fix, and the [#176](https://github.com/saveenr/VisioAutomation/issues/176) `<frameworkAssembly>` removal (all next-NuGet-release pickup); [`VisioPowerShell/CHANGELOG.md`](VisioAutomation_2010/VisioPowerShell/CHANGELOG.md) `[Unreleased]` has the SevenPens Author entry plus the [#177](https://github.com/saveenr/VisioAutomation/issues/177) Release-builds entry (next-PS-module-release pickup).

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

**Previous session (2026-05-08) summary:** Single-theme session, focused on test naming and convention rollout. Six issues closed in sequence:

1. **[#154](https://github.com/saveenr/VisioAutomation/issues/154)** &mdash; test-coverage gap audit. Output: [`docs/futures/test-coverage-gaps.md`](docs/futures/test-coverage-gaps.md), a prioritized gap list across the public API surface (~340 types) cross-referenced against the test inventory (211 enumerated). Headline finding: PowerShell cmdlet coverage is the biggest gap (~56 of 70 cmdlets without dedicated tests, exactly the regression class that shipped in module 4.6.1).
2. **[#165](https://github.com/saveenr/VisioAutomation/issues/165)** &mdash; worst-offender renames + forward-looking naming convention codified in [`docs/TESTING.md`](docs/TESTING.md#naming-conventions). 36 method renames + 3 class renames + 3 file renames covering numbered-suffix series, redundant `Test_` infix, mis-labeled singletons.
3. **[#166](https://github.com/saveenr/VisioAutomation/issues/166)** &mdash; class-redundant prefix sweep + file-name underscore cleanup across VTest.Scripting and VTest.Models. ~80 method prefix drops + 30 file/class renames.
4. **[#167](https://github.com/saveenr/VisioAutomation/issues/167)** &mdash; VTest.PowerShell sweep, vague-name improvements, stylistic outliers, helper-method visibility tightening. ~21 helpers made private; 15 vague names improved with body-inspected `Subject_Scenario_ExpectedOutcome` names.
5. **[#168](https://github.com/saveenr/VisioAutomation/issues/168)** &mdash; "Scenarios" kitchen-sink splits. Four kitchen-sink tests split into 13 (Hyperlinks, Controls, Selection, CustomProps); three mis-labeled singletons just renamed (Undo, CloseDocument, ConnectionPoints).
6. **[#169](https://github.com/saveenr/VisioAutomation/issues/169)** &mdash; final pass: VTest project prefix sweep + AddRemove paired-action splits. 22 prefix drops + 3 AddRemove tests split into 9.

After all the rename work, manually ran the full suite for the first time today; surfaced two pre-existing `doc.Close()` (no `force=true`) bugs in [`DirectedGraphDrawModelTests.cs`](VisioAutomation_2010/VTest.Models/DirectedGraphDrawModelTests.cs) (Drawing9 and Drawing20 save prompts that needed manual `Don't Save` clicks). Fixed in `98753108` by switching the 5 occurrences to `doc.Close(true)` and adding `using VisioAutomation.Extensions;`.

Test count: 211 enumerated &rarr; 226 enumerated (230 runtime). 8 commits to master across this session (`5946629a` &rarr; `98753108`).

Plus filed [#170](https://github.com/saveenr/VisioAutomation/issues/170) for next session focus (LINQ provider for ShapeSheets).

**This session (2026-05-08b) summary:** Long working session, three major implementation threads plus broad backlog hygiene.

1. **[#170](https://github.com/saveenr/VisioAutomation/issues/170) LINQ for ShapeSheet &mdash; motivation doc, spike deferred.** Rather than spike the implementation, drafted [`docs/futures/linq-shapesheet-before-after.md`](docs/futures/linq-shapesheet-before-after.md) (~250 lines) showing 9 ShapeSheet-query scenarios in today's API vs. a hypothetical LINQ shape, with an honest verdict per scenario (where LINQ wins, where it's a wash, where it may not fit). Doc closes with 7 design questions for whoever picks the spike up. Doc lives on the `experiment/linq-shapesheet` branch (commit `ac422911`, pushed to origin); spike code itself was deferred. **Issue [#170](https://github.com/saveenr/VisioAutomation/issues/170) rescoped from "LINQ provider for ShapeSheet queries" to "ShapeSheet query ergonomics (LINQ and other shapes)" and moved CY26Q2 &rarr; CY27Q1** at user request, since the user wanted to keep thinking about it (may yield non-LINQ ideas).

2. **[#152](https://github.com/saveenr/VisioAutomation/issues/152) Phase H1 closed.** Owner status calls confirmed all 8 surviving sibling repos. [`docs/RELATED-REPOS.md`](docs/RELATED-REPOS.md) updated to drop "(preliminary)" tags, reflect Visio-Font-Compare reclassification (abandoned &rarr; paused), remove the deleted `visio-templates` row, and record CY27 work links per row. Per-repo `README.md` updates pushed to all 8 sibling repos with consistent Status blockquotes. `visio-templates` deleted (by user via UI; gh CLI auth lacked `delete_repo` scope). #152 closed with summary comment. Surfaced `gh auth refresh -h github.com -s delete_repo` as the way to grant the scope if ever needed in future.

3. **[#176](https://github.com/saveenr/VisioAutomation/issues/176) + [#177](https://github.com/saveenr/VisioAutomation/issues/177) implemented and closed.**
   - **[#176](https://github.com/saveenr/VisioAutomation/issues/176)** (nuspec `<frameworkAssembly>` trim): removed the redundant declaration from [`NuGet/VisioAutomation2010.nuspec`](NuGet/VisioAutomation2010.nuspec); CHANGELOG `[Unreleased]` Removed entry. Closed via commit `47d3ded6`.
   - **[#177](https://github.com/saveenr/VisioAutomation/issues/177)** (PS module Release builds): parameterized [`InstallForCurrentUser.ps1`](VisioAutomation_2010/VisioPowerShell/InstallForCurrentUser.ps1) with `-Configuration` (default `Debug` for dev convenience), updated [`Publish-VisioPSToGallery.ps1`](VisioAutomation_2010/VisioPowerShell/Publish-VisioPSToGallery.ps1) to pass `-Configuration Release`, switched [`release-psmodule.yml`](.github/workflows/release-psmodule.yml) CI build + stage to Release. CHANGELOG `[Unreleased]` Changed entry. Closed via commit `b8710d96`. Filed [#181](https://github.com/saveenr/VisioAutomation/issues/181) as the NuGet-side follow-up (release-nuget.yml + nuspec).

4. **Filed 11 tracking issues for backlog items previously without GitHub issues.** Distributed across CY26Q3 (5: [#173](https://github.com/saveenr/VisioAutomation/issues/173), [#176](https://github.com/saveenr/VisioAutomation/issues/176), [#177](https://github.com/saveenr/VisioAutomation/issues/177), [#178](https://github.com/saveenr/VisioAutomation/issues/178), [#181](https://github.com/saveenr/VisioAutomation/issues/181)), CY26Q4 (2: [#172](https://github.com/saveenr/VisioAutomation/issues/172), [#179](https://github.com/saveenr/VisioAutomation/issues/179)), CY27Q1 (2: [#171](https://github.com/saveenr/VisioAutomation/issues/171), [#174](https://github.com/saveenr/VisioAutomation/issues/174)), CY27Q2 (2: [#175](https://github.com/saveenr/VisioAutomation/issues/175), [#180](https://github.com/saveenr/VisioAutomation/issues/180)). Plus 6 per-repo CY27 issues filed on each sibling repo (VisioAutomation.VDX#5, Visio-Power-Tools#2, Visio-Export-Pages-To-Docs#5, Visio-Code-Samples#1, visio-reference#1, Visio-Font-Compare#1) capturing planned CY27 hygiene/refresh work. [`docs/MILESTONES.md`](docs/MILESTONES.md) tables filled in with the new issue numbers.

5. **Doc cleanup pass.** Pruned 4 stale entries from [`docs/futures/docs.md`](docs/futures/docs.md) (Tier 3 done in prior session; version-compat done in prior session; custom-properties matrix that [#145](https://github.com/saveenr/VisioAutomation/issues/145) closed earlier; refresh-resources / annotate-VS2026-note both completed this session). Updated [`docs/FUTURES.md`](docs/FUTURES.md) index to match. Removed the line-150 "Tier 3 .NET-side coverage" pending mention from [`docs/MILESTONES.md`](docs/MILESTONES.md). Fixed 2 cross-refs in `docs/futures/docs.md` that pointed to removed entries.

6. **Gitbook updates** (.NET only). [`compiling.md`](https://saveenr.gitbook.io/visioautomation/compiling): VS 2026 note now links to [#171](https://github.com/saveenr/VisioAutomation/issues/171) + carries Last verified date; **separately**, fixed a stale dev-pack requirement (the chocolatey install line was wrong since the SDK migration's NuGet ref-assemblies switch). [`resources/README.md`](https://saveenr.gitbook.io/visioautomation/resources): added Microsoft Learn entry, added in-repo ARCHITECTURE/GLOSSARY refs, demoted but kept the 2003 book.

9 commits to master this session (`e4af7642` &rarr; `b8710d96`). 2 commits to .NET gitbook. 8 commits across the 8 sibling repos (one README each). 11 issues filed in main repo + 6 in sibling repos; 3 closed ([#152](https://github.com/saveenr/VisioAutomation/issues/152), [#176](https://github.com/saveenr/VisioAutomation/issues/176), [#177](https://github.com/saveenr/VisioAutomation/issues/177)).

## Next session priorities

**Start next session with [#156](https://github.com/saveenr/VisioAutomation/issues/156) &mdash; Decide: is `VisioScripting` part of the project's promised public API, or an internal that shouldn't be relied on?** Decision-only issue (S effort) but the biggest single unblocker in the backlog: answering it directly unblocks [#131](https://github.com/saveenr/VisioAutomation/issues/131) (the largest remaining doc-write &mdash; `VisioScripting.Client` undocumented on gitbook) and shapes the strategy under [#172](https://github.com/saveenr/VisioAutomation/issues/172) (revise user-facing docs).

**Background to load:**

- [`docs/futures/docs.md`](docs/futures/docs.md), entry "Decide whether to document `VisioScripting` as a public API" &mdash; framing of the question, current docs state (just `Get-VisioClient` + two technical-notes pages), surface-size estimate (~15 helper/command-set pages, similar to the .NET-side Tier 1+2+4 work already shipped).
- The codebase: [`VisioAutomation_2010/VisioScripting/`](VisioAutomation_2010/VisioScripting/) &mdash; particularly the `Client` object grouping commands by topic (`Document`, `Page`, `Selection`, `Shape`, `ShapeSheet`, `CustomProperty`, etc., ~22 groups).
- The PS-side escape hatches today: `cmdlets/other-cmdlets.md`'s `Get-VisioClient`; `technical-notes/getting-the-current-scriptingsession.md`; `technical-notes/use-visioautomation.md`.
- 2026-05-05 doc-review feedback noted in `docs/futures/docs.md`: the source [`readme.md`](readme.md)'s quick-start opens with a `VisioScripting.Client` snippet, so a new C# reader's first impression is an undocumented type. Strongest external argument for "yes, public."

**Decision shape &mdash; three plausible answers:**

1. **Yes, public API.** Document the full surface (~15 pages); commit to API stability; no more "shifts to suit cmdlet needs" without thinking about external consumers.
2. **No, internal.** Leave undocumented; fix `readme.md`'s quick-start to use `VisioAutomation` directly; close [#131](https://github.com/saveenr/VisioAutomation/issues/131) as won't-fix; mark types `[EditorBrowsable(Never)]` or similar.
3. **Hybrid.** Document `Client.<Group>.<Method>(...)` *interface* as public-stable; leave implementation classes (`*Commands`, `*Loaders`) undocumented as still-mutable internals.

Each answer changes the work scope for [#131](https://github.com/saveenr/VisioAutomation/issues/131) and shapes [#157](https://github.com/saveenr/VisioAutomation/issues/157) (long-term docs location). It also touches [#162](https://github.com/saveenr/VisioAutomation/issues/162) (PSVA Phase A scoping) and [#164](https://github.com/saveenr/VisioAutomation/issues/164) (VisioCmdlet base class) since both can shift the cmdlet/Scripting boundary.

**Acceptance:** comment on [#156](https://github.com/saveenr/VisioAutomation/issues/156) with the decision, one-paragraph rationale, and a brief "implications for #131, #172, #157" bullet. Close [#156](https://github.com/saveenr/VisioAutomation/issues/156).

**Lower priority same semester** (CY26Q2 is in close-out mode):

- **[#151](https://github.com/saveenr/VisioAutomation/issues/151) triage tracker** &mdash; calendar-bound. Revisit ~2026-05-27: if `@tcox8` hasn't confirmed [Visio PS 4.7.0](https://github.com/saveenr/VisioAutomation/releases/tag/VisioPS_4.7.0) fixed their case by then, close [#117](https://github.com/saveenr/VisioAutomation/issues/117) as fixed-by-#144 then close [#151](https://github.com/saveenr/VisioAutomation/issues/151).

**Quick-pickup follow-ups if bandwidth permits** (all CY26Q3, all S&ndash;M effort):

- **[#178](https://github.com/saveenr/VisioAutomation/issues/178)** finish `Visio.psd1` deprecation cleanup (RootModule rename + PowerShellVersion bump). Customer-impact analysis pre-done; one-line edits.
- **[#181](https://github.com/saveenr/VisioAutomation/issues/181)** mirror [#177](https://github.com/saveenr/VisioAutomation/issues/177) for the NuGet release flow (release-nuget.yml + nuspec). Same pattern, same effort.
- **[#173](https://github.com/saveenr/VisioAutomation/issues/173)** cmdlet-binding test first slice. Closes the regression class that shipped 4.6.1's four bugs.

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
