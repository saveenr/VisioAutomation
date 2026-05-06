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

**Visio PowerShell 4.7.0 shipped to PSGallery on 2026-05-06** (tag `VisioPS_4.7.0`, first VisioPS [GitHub Release](https://github.com/saveenr/VisioAutomation/releases/tag/VisioPS_4.7.0)). The release closes [#144](https://github.com/saveenr/VisioAutomation/issues/144) and is the natural completion of the long-running thread that started with [#117](https://github.com/saveenr/VisioAutomation/issues/117) (reporter pinged with the one-line fix).

**Issue #144 — CustomPropertyCells encoding ergonomics — done.** The fix landed across three coordinated pieces (committed as `da962fe5`):

- **Typed setters** on `CustomPropertyCells`: `SetString` / `SetNumber` / `SetBool` / `SetDate` / `SetFormula`. Same idea on `UserDefinedCellCells`: `SetString` / `SetFormula`. Each writes a correctly-encoded Visio formula and (where applicable) sets `Type` to match.
- **`Value` → `Formula` rename** on both classes (with `Value` kept as an `[Obsolete]` source-compat alias).
- **F detect-and-rethrow** in `CustomPropertyHelper.Set` and `UserDefinedCellHelper.Set`: Visio's opaque `COMException: #NAME?` is now wrapped in an `ArgumentException` with a self-explanatory diagnostic pointing at the typed setters. Original `COMException` preserved as `InnerException`.
- **Cross-layer migration** in the same PR: VisioPowerShell `Set-VisioCustomProperty` cmdlet, `DGShapeInfo` (XML loader path), and the `VSamples.Docs/SetCustomProperties.cs` doc-snippet all switched to the typed setters. The `EncodeValues()` backstop is preserved in `VisioScripting.CustomPropertyCommands` for `-Value`-form callers.

Engineering reference at [`docs/internal/custom-property-encoding.md`](docs/internal/custom-property-encoding.md): full per-Type behavior matrix locked in by 5 + 2 round-trip MSTest methods in [`VTest/Core/Shapes/`](VisioAutomation_2010/VTest/Core/Shapes/). User-facing matrix added to the [Custom properties gitbook page](https://saveenr.gitbook.io/visioautomation/custom-properties); cross-linked from the PS-side [Directed graphs from code](https://saveenr.gitbook.io/visiopowershell/automatic-diagrams/drawing-directed-graphs) page. Both gitbook repos pushed. [#145](https://github.com/saveenr/VisioAutomation/issues/145) (extend gitbook with matrix) was filed and shipped in this session — ready to close.

**Pre-existing fix landed as a drive-by:** `VerifyFormulaLiteralSize` was a 32-bit-only `Marshal.SizeOf(CellValue)` assertion that broke when run under the 64-bit testhost. Replaced the hardcoded `4` with `System.IntPtr.Size` (`5ae458f2`).

**Repo state:** master at `8438eb66`, working tree clean. PSGallery: `Visio` 4.7.0 live. NuGet: `VisioAutomation2010` still 2.6.0 (no NuGet release in this session). `[Unreleased]` in [`NuGet/CHANGELOG.md`](NuGet/CHANGELOG.md) has accumulated entries from #144 + the open-issues backlog pass that will go into the next NuGet release. Both gitbook repos in sync with origin.

**Test infrastructure:** all four projects green — VTest 101/101, VTest.Models 59/59, VTest.Scripting 34/34, VTest.PowerShell 17/17. Total 211. (Up from 177 baseline; the increase comes from the new characterization + setter tests added in this session, plus the prior open-issues-backlog pass that added PS tests.)

**Dev-environment note:** Smart App Control on Windows 11 can transiently block freshly-built test DLLs (`Application Control policy has blocked this file`, HRESULT 0x800711C7, Code Integrity policy ID `{0283ac0f-fff1-49ae-ada1-8a933130cad6}`). The block clears once SAC's cloud reputation cache catches up (30-60s); VS 2022 Test Explorer hits the same mechanic. Durable fix: turn Smart App Control off via Settings → Privacy & security → Windows Security → App & browser control → Smart App Control settings → Off. **One-way operation** — re-enabling requires reinstalling Windows.

**Prior session context** (still recent, still relevant):
- **Phase 3 SDK migration done.** All 11 csprojs SDK-style + PackageReference + Central Package Management; dev-pack install requirement gone via reference-assemblies packages; CI build ~70s. Detail in [`docs/COMPLETED.md`](docs/COMPLETED.md#phase-3--modernization-in-progress).
- **Open-issues backlog pass** closed nine issues (#105 / #128 / #129 / #130 / #138 / #139 / #140 / #141 / #142 / #143) and shipped the directed-graph XML attribute work (`connectortype`, `direction`, `layout`).
- **Visio PowerShell 4.6.1 shipped 2026-05-03** with four cmdlet bug fixes; the 4.7.0 release in this session bundles everything that landed since.

**Next session priorities:**

1. **Phase 2 prep — cut the final `VisioAutomation2010` NuGet release.** `[Unreleased]` in [`NuGet/CHANGELOG.md`](NuGet/CHANGELOG.md) has accumulated substantial entries (typed setters, `Formula` rename, `ArgumentException` wrapping, plus directed-graph XML attribute work). Discussion-first: the deferred version-policy question (NuGet at `2.6.x` vs PS module now at `4.7.0`: converge or stay divergent?) needs an answer before pulling the trigger.
2. **Release CI completion** — wrap the manual flow that just shipped 4.7.0 into reusable workflows. The 4.7.0 PSGallery publish was via [`Publish-VisioPSToGallery.ps1`](VisioAutomation_2010/VisioPowerShell/Publish-VisioPSToGallery.ps1); the GitHub Release was created manually via `gh release create`. The two flows currently conflict on tag-creation, so wrapping them needs a small refactor (e.g. teach the publish script to skip tagging if the tag already exists, or move tag-creation out of both into a third script). Plan in [`docs/futures/releases.md`](docs/futures/releases.md#automate-releases-via-github-ci-in-progress).
3. **Small follow-ups deferred from the migration** — the `<frameworkAssembly assemblyName="Microsoft.Office.Interop.Visio" />` line in [`NuGet/VisioAutomation2010.nuspec`](NuGet/VisioAutomation2010.nuspec) is now redundant with the bundled-PIA-via-NuGet behavior (PIA package's `lib/` ships in the .nupkg). Trim as a small standalone commit when convenient.
4. **(Post-2026-10-13) TFM bump** — only after Windows 10 LTSB 2016 leaves Extended Support. See [enterprise compat memory](../../.claude/projects/C--Users-savee-Documents-GitHub-VisioAutomation/memory/enterprise_compat_ltsb2016.md). Unblocks VS 2026 move.

## Other docs in this repo

- [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md) — projects, dependencies, central concepts
- [docs/BUILDING.md](docs/BUILDING.md) — full build / test / install reference (incl. dev-pack install commands)
- [docs/TESTING.md](docs/TESTING.md) — test-suite design and conventions (shared `Framework.VTest` base, per-testhost Visio singleton, `[AssemblyCleanup]` orphan-prevention, MSTEST0030 enforcement). Per-project READMEs sit next to each test csproj.
- [docs/GLOSSARY.md](docs/GLOSSARY.md) — Visio and codebase terminology
- [docs/ROADMAP.md](docs/ROADMAP.md) — staged plan (Phase 1 / 2 / 3); read first for orientation
- [docs/FUTURES.md](docs/FUTURES.md) — index of backlog items, split by topic into [`docs/futures/*.md`](docs/futures/)
- [docs/OVERVIEW.md](docs/OVERVIEW.md) — entry point, links to the above
