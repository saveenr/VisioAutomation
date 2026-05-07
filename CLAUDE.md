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

**Repo state:** master at HEAD (`77a0268c`). Working trees clean across all three repos (this repo + the two gitbook repos). PSGallery: `Visio` 4.7.2 live. NuGet: `VisioAutomation2010` 3.0.0 live. [`NuGet/CHANGELOG.md`](NuGet/CHANGELOG.md) `[Unreleased]` has the SevenPens brand-swap entry plus the [#82](https://github.com/saveenr/VisioAutomation/issues/82) `MsaglRenderer.format_shape` fix (next NuGet release pickup); [`VisioPowerShell/CHANGELOG.md`](VisioAutomation_2010/VisioPowerShell/CHANGELOG.md) `[Unreleased]` is empty.

**Test infrastructure:** all four projects green &mdash; VTest 101/101, VTest.Models **60/60**, VTest.Scripting 34/34, VTest.PowerShell 20/20. Total **215 tests** (added `DirectedGraph_NodeSizeIsHonored_WhenCellsAlsoSet` for the [#82](https://github.com/saveenr/VisioAutomation/issues/82) regression). Each test run still leaves zero Visio orphan processes.

**Last session (2026-05-07) summary:** 7 issues closed ([#80](https://github.com/saveenr/VisioAutomation/issues/80), [#82](https://github.com/saveenr/VisioAutomation/issues/82), [#102](https://github.com/saveenr/VisioAutomation/issues/102), [#105](https://github.com/saveenr/VisioAutomation/issues/105), [#149](https://github.com/saveenr/VisioAutomation/issues/149), [#150](https://github.com/saveenr/VisioAutomation/issues/150), [#153](https://github.com/saveenr/VisioAutomation/issues/153)), 3 new filed ([#161](https://github.com/saveenr/VisioAutomation/issues/161) version-compat tables, [#162](https://github.com/saveenr/VisioAutomation/issues/162) Phase A scoping, [#163](https://github.com/saveenr/VisioAutomation/issues/163) Phase B implementation), 6 commits to master, 1 follow-up still pending ([#117](https://github.com/saveenr/VisioAutomation/issues/117) revisit ~2026-05-27 then [#151](https://github.com/saveenr/VisioAutomation/issues/151) closes). New ADR pattern: [`docs/decisions/`](docs/decisions/) folder established with [`tests-need-visio.md`](docs/decisions/tests-need-visio.md) as the inaugural entry. PSVA gap audit landed in [`docs/futures/build-and-code.md`](docs/futures/build-and-code.md).

**Dev-environment note:** Smart App Control on Windows 11 can transiently block freshly-built test DLLs (HRESULT 0x800711C7). The block clears in 30&ndash;60s as SAC's cloud reputation cache catches up. Durable fix is to turn SAC off (one-way operation; can't re-enable without reinstalling Windows). CI runners are not subject to this.

## Next session priorities (CY26Q2)

Four [CY26Q2 GitHub milestone](https://github.com/saveenr/VisioAutomation/milestone/2) items still open. Due 2026-06-30. **Suggested tackle order for next session:**

1. **[#163 &mdash; Add `ValueFromPipeline` parameter set to `Connect-VisioShape`](https://github.com/saveenr/VisioAutomation/issues/163)**: ~1&ndash;2 hours, smallest remaining item. Back-compat-safe Phase B from the [#150](https://github.com/saveenr/VisioAutomation/issues/150) PSVA gap audit. Issue body has design questions to settle first (parameter-set shape, single-vs-array on the pipeline side, Begin/Process/End batching strategy) &mdash; *spend the first 15 min on those decisions*, then implement + test + CHANGELOG. Touches [`ConnectVisioShape.cs`](VisioAutomation_2010/VisioPowerShell/Commands/VisioShape/ConnectVisioShape.cs); a regression test goes in [`VTest.PowerShell`](VisioAutomation_2010/VTest.PowerShell/). Closing this clears all of Milestone D's CY26Q2 work.

2. **[#161 &mdash; Version compatibility tables on user-facing gitbooks](https://github.com/saveenr/VisioAutomation/issues/161)**: ~1 day. Two new gitbook pages, one per product (`VisioAutomation_GitBook_Docs` for the NuGet matrix, `VisioPowerShellDocs` on the `visiops_v4_docs` branch for the VisioPS matrix). Per-row data needs sourcing from git tags + csproj `TargetFramework` + `Visio.psd1` `PowerShellVersion` + `Directory.Packages.props` &mdash; no guessing. Best as a focused single-session block with both gitbook checkouts open simultaneously; cross-link the pages and link from `readme.md` + both `CHANGELOG.md`.

**Lower priority for the same semester:**

- **[#152](https://github.com/saveenr/VisioAutomation/issues/152) Visio repo portfolio audit Phase H1**: centralized index + per-repo READMEs across the 9 sibling Visio repos. Bounded but biggest of the open four; Phase H2 (per-repo retire/merge decisions) is its own follow-up. Defer to whenever the smaller items are clear.
- **[#151](https://github.com/saveenr/VisioAutomation/issues/151) triage tracker**: open intentionally as the home for the [#117](https://github.com/saveenr/VisioAutomation/issues/117) revisit. **Calendar reminder: ~2026-05-27** &mdash; if `@tcox8` hasn't confirmed [Visio PS 4.7.0](https://github.com/saveenr/VisioAutomation/releases/tag/VisioPS_4.7.0) fixed their case by then, close [#117](https://github.com/saveenr/VisioAutomation/issues/117) as fixed-by-#144, then close [#151](https://github.com/saveenr/VisioAutomation/issues/151).

**Out of scope unless we finish CY26Q2 with bandwidth to spare:** anything in [CY26Q3](https://github.com/saveenr/VisioAutomation/milestone/3) (which now includes [#162](https://github.com/saveenr/VisioAutomation/issues/162) Phase A scoping, the natural follow-on to [#163](https://github.com/saveenr/VisioAutomation/issues/163)) or later semesters.

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
