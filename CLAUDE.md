# CLAUDE.md

Project-specific guidance for Claude Code sessions in this repo. Loaded automatically.

## What this is

A .NET Framework library plus a PowerShell module that automate Microsoft Visio via COM interop. Full picture: [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md).

## The 2026 refresh — read this before making changes

Active branch: `master`. Phase 1 of the refresh has merged; Phase 2 (the final release of the `VisioAutomation2010` NuGet) and Phase 3 (modernization) are still ahead. Work is staged in three phases per [docs/FUTURES.md](docs/FUTURES.md):

1. **Phase 1 — VS 2022 cleanup** *(done; merged to master 2026-05-03)*. Stayed on the then-current TFMs (.NET Framework 4.5.2 for shipping libs, .NET Framework 4.7.2 for tests) and shipped the **Visio PowerShell 4.6.1** release as the conclusion of this phase.
2. **Phase 2 — Cut a final release** of the `VisioAutomation2010` NuGet (currently `2.6.0`) with refreshed docs. Two prereqs are deferred and need user discussion: the version-number policy (NuGet `2.6.0` vs PS module `4.6.1`) and a leftover-Visio-process flakiness investigation.
3. **Phase 3 — Modernization.** Move to VS 2026 (which requires bumping TFMs to 4.7.2), modern C#, possibly modern .NET, automated releases.

When in doubt whether a change fits the current phase, check [docs/FUTURES.md](docs/FUTURES.md).

## Build prerequisites

The shipping libs target .NET Framework 4.5.2. Modern Windows install media don't include the v4.5.2 reference assemblies, so **the .NET Framework 4.5.2 Developer Pack must be installed** on every dev machine and CI runner. Microsoft does not publish a winget manifest for 4.5.x — use chocolatey (`choco install netfx-4.5.2-devpack -y`) or the Microsoft download page. Full instructions in [docs/BUILDING.md](docs/BUILDING.md).

If you see `MSB3644: The reference assemblies for .NETFramework,Version=v4.5.2 were not found` — that's the dev pack not being installed.

## Build commands (verified working)

```bash
MSBUILD="/c/Program Files/Microsoft Visual Studio/2022/Community/MSBuild/Current/Bin/MSBuild.exe"
"$MSBUILD" VisioAutomation_2010/VisioAutomation2010.sln -t:Restore -p:RestorePackagesConfig=true
"$MSBUILD" VisioAutomation_2010/VisioAutomation2010.sln -p:Configuration=Debug -m
```

- **Do not** use `dotnet build` — projects are legacy csproj + packages.config.
- **Do not** use VS 2026's MSBuild (under `Program Files\Microsoft Visual Studio\18\`) — its .NET Framework floor is 4.6.2 and the shipping libs are on 4.5.2. This is a Phase 3 unblocker.

Full reference (IDE flow, test invocation, dev-pack install commands): [docs/BUILDING.md](docs/BUILDING.md).

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

**Doc audit:** complete on both gitbook repos. Reader-facing summaries at [VisioPowerShellDocs/documentation-changes.md](https://saveenr.gitbook.io/visiopowershell/documentation-changes) and [VisioAutomation_GitBook_Docs/documentation-changes.md](https://saveenr.gitbook.io/visioautomation/documentation-changes). Every cmdlet has a complete page in the standard `## Syntax` / `## Parameters` / `## Examples` / `## See also` layout, and the .NET-side helper-class coverage was extended in three tiers (Tier 1: Hyperlinks / Lock cells / Control handles / Connection points / Connectors. Tier 2: Shape format-layout-xform / Page cells / Text formatting / Geometry / Application. Tier 4: Analyzers / Visio error log / UndoScope / Exceptions / full Extension-methods rewrite). Only Tier 3 (the `VisioAutomation.Models` project) is still pending and is recorded in `docs/FUTURES.md` under *"Expand .NET-side doc coverage"*.

**Repo state:** all three repos in sync with origin, working trees clean. No in-flight branches.

**Test infrastructure:** all 163 tests across the four test projects pass as of 2026-05-04 (VTest 94, VTest.Models 37, VTest.Scripting 28, VTest.PowerShell 4). Fixes that landed to get there:

- `12027821` &mdash; removed the legacy MSTest v1 project-type GUID (`{3AC096D0-…}`) from all four test csprojs (was making VS Test Explorer use the legacy discovery path). Added the `System.Threading.Tasks.Extensions 4.5.4` package as a transitive dep of MSTest's runner.
- `5606adcc` &mdash; enabled `<AutoGenerateBindingRedirects>` + `<GenerateBindingRedirectsOutputType>` so library projects emit `.dll.config` files with binding redirects.
- `5cbf11cd` &mdash; removed redundant `[DeploymentItem]` attributes (8 of them across `XmlErrorLogTests`, `DrawModel_DirectedGraph`, `DrawModel_OrgChartTests`). The data files are already `CopyToOutputDirectory=Always`, so the attributes were redundant; their only effect was triggering VS Test Explorer's deployment-mode behavior, which dropped runtime dependencies on the floor.
- `fb1799d4` &mdash; reverted an attempt to use a solution-level `default.runsettings` with `DeploymentEnabled=false` (it broke previously-passing tests in VS Test Explorer; not the right approach).
- The last failing test was `Dom_DrawOrgChart` &mdash; fixed by version-guarding the template filename (`orgchart.vst` for Visio &lt; v15, `orgch_u.vstx` for Visio 2013+), parallel to the already-version-guarded master name on the line above. Visio 2013 replaced binary `.vst` templates with XML-based `.vstx` and the user's M365 install only ships `.vstx`. **Note:** [`OrgChartStyling.cs:9`](VisioAutomation_2010/VisioAutomation.Models/Documents/OrgCharts/OrgChartStyling.cs:9) has the same bug pattern in production code (`Visio2013Template = "orgch_u.vst"`); flagged as a follow-up before the Phase 2 NuGet release.

## Resume here: next session is test cleanup, then release CI

**Phase 1 of the next session — the broader test-cleanup pass.** All tests pass; the structural cleanup tracked in [`docs/FUTURES.md`](docs/FUTURES.md) under *"General cleanup of the test projects"* is the natural next chunk: deterministic setup/teardown, robustness against leftover Visio processes, the `MSB3270` x86/AnyCPU mismatch on `VTest.PowerShell` (visible in the build log), modernization, per-project READMEs explaining what each project covers.

**Phase 2 of the next session — GitHub Actions release CI.** Only after tests are clean. Three deliverables, full plan in [`docs/FUTURES.md`](docs/FUTURES.md) &rarr; *"Automate releases via GitHub CI"*:

1. **PSGallery publish** of the `Visio` PowerShell module (canonical script already exists: [`Publish-VisioPSToGallery.ps1`](VisioAutomation_2010/VisioPowerShell/Publish-VisioPSToGallery.ps1)).
2. **nuget.org publish** of the `VisioAutomation2010` NuGet package (no canonical script yet; metadata in [`NuGet/VisioAutomation2010.nuspec`](NuGet/VisioAutomation2010.nuspec) currently `2.6.0`).
3. **GitHub Releases** with the built binaries / `.nupkg` / module zip attached as downloadable artifacts.

After CI lands, the other Phase-2-deferred items (version-number policy, leftover-Visio-process flakiness investigation) become the next conversation.

## Other docs in this repo

- [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md) — projects, dependencies, central concepts
- [docs/BUILDING.md](docs/BUILDING.md) — full build / test / install reference (incl. dev-pack install commands)
- [docs/GLOSSARY.md](docs/GLOSSARY.md) — Visio and codebase terminology
- [docs/FUTURES.md](docs/FUTURES.md) — staged backlog of cleanup/modernization work
- [docs/OVERVIEW.md](docs/OVERVIEW.md) — entry point, links to the above
