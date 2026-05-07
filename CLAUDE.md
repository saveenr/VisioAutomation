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

**VisioAutomation NuGet 3.0.0 shipped to nuget.org on 2026-05-07** ([tag `VisioAutomation_3.0.0`](https://github.com/saveenr/VisioAutomation/releases/tag/VisioAutomation_3.0.0); [package page](https://www.nuget.org/packages/VisioAutomation2010/3.0.0)). First NuGet release published end-to-end via the new CI flow (the freshly-shipped `release-nuget.yml` + `publish-nuget.yml`). Phase 2 of the 2026 refresh is now complete on the NuGet side. Major version bump per SemVer driven by four flagged behavior changes in the rolled changelog: `ArgumentException`-instead-of-`COMException` for malformed property values, `DrawDirectedGraphDocument` honoring per-page `UseDynamicConnectors`, stricter `<directedgraph>` root validation, and net40 &rarr; net452 minimum TFM. Source compatibility for property access preserved via the `[Obsolete]` `Value` &rarr; `Formula` alias from #144.

**Microsoft-package compliance gate quirk discovered and worked around** (operational reference, [docs/futures/releases.md](docs/futures/releases.md#microsoft-package-compliance-gate-on-the-saveenr-nugetorg-account-operational-quirk-discovered-2026-05-07-during-the-300-publish)). nuget.org rejected the first 3.0.0 publish attempt under the `saveenr` API key with HTTP 400, citing Microsoft-package compliance violations on `<authors>` and `<copyright>`. Root cause: the `saveenr` nuget.org account is classified as Microsoft-affiliated (likely from this project's Visio-team origin); compliance enforcement was tightened since 2.6.0 and now mandates exact-string `<authors>Microsoft</authors>` + `<copyright>&copy; Microsoft Corporation</copyright>` on uploads from classified accounts. Workaround: [`SevenPens`](https://www.nuget.org/profiles/SevenPens) (non-affiliated) was added as a co-owner of `VisioAutomation2010`, and `NUGET_API_KEY` was rotated to a SevenPens-generated key. Same `.nupkg` content (unchanged metadata) then pushed successfully on the second attempt. **The gate fires on uploader's account classification, not on owner-list classification** &mdash; so saveenr remains a co-owner without consequence; only the API key identity matters.

**[`publish-nuget.yml`](.github/workflows/publish-nuget.yml) shipped 2026-05-07** &mdash; sibling to [`publish-psmodule.yml`](.github/workflows/publish-psmodule.yml), completing the release-CI plan in [`docs/futures/releases.md`](docs/futures/releases.md#automate-releases-via-github-ci-in-progress). Both deliverables (PS module + NuGet package) now have full two-step CI: `release-*.yml` builds + creates the GH Release with the artifact attached, then `publish-*.yml` downloads that artifact and pushes to the public feed. Same shape across the board: `workflow_dispatch`-triggered, takes the release tag as input, verifies the staged artifact's embedded version metadata against the tag before pushing, has a `dry_run` switch, idempotent on re-runs (`--skip-duplicate` for `dotnet nuget push`; the PS-module side already had the equivalent semantics), and verifies the upload landed via the feed's index API with retry. The NuGet half adapts the pattern with: tag prefix `VisioAutomation_` (vs. `VisioPS_`), nupkg-as-zip nuspec verification (vs. `Visio.psd1` `Import-PowerShellDataFile`), nuget.org V3 flat-container poll with a longer ~5 min retry window (vs. PSGallery's ~30 s &mdash; nuget.org's indexing lags 1-3 min and full validation can run 15+ min). No equivalent of `publish-psmodule.yml`'s TLS-1.2 step or `CmdletsToExport` drift check is needed (PS 7 / `dotnet` handle TLS, and NuGet has no per-cmdlet export manifest).

**Visio PowerShell 4.7.2 shipped to PSGallery on 2026-05-07** ([tag `VisioPS_4.7.2`](https://github.com/saveenr/VisioAutomation/releases/tag/VisioPS_4.7.2)). First release published end-to-end via the new CI flow ([`release-psmodule.yml`](.github/workflows/release-psmodule.yml) builds + creates the GH Release with zip; [`publish-psmodule.yml`](.github/workflows/publish-psmodule.yml) downloads the zip and pushes to PSGallery). 4.7.2 contains a single user-visible fix &mdash; `Set-VisioUserDefinedCell -Value` / `-Prompt` now actually work for plain string arguments &mdash; plus the `CmdletsToExport='*'` &rarr; explicit-list manifest hygiene change.

**The CI release flow is now fully operational.** Both halves landed this session:

- New [`publish-psmodule.yml`](.github/workflows/publish-psmodule.yml) workflow (`workflow_dispatch`-triggered, takes the release tag as input, downloads the GH Release zip, verifies `ModuleVersion` and `CmdletsToExport` drift, force-TLS-1.2, runs `Publish-Module`, verifies via `Find-Module` with retry, has a `dry_run` option that exercises everything except the upload).
- [`Publish-VisioPSToGallery.ps1`](VisioAutomation_2010/VisioPowerShell/Publish-VisioPSToGallery.ps1)'s tag step refactored to be idempotent so the local script and the workflow coexist for fallback / out-of-band publishes.
- **Latent bug surfaced and fixed in both `release-psmodule.yml` and `release-nuget.yml`**: PowerShell here-strings (`@'...'@`) inside YAML literal block scalars (`run: |`) are mutually exclusive due to indentation rules. The YAML had been parsing as broken since 2026-05-04 (43 phantom 0-second failed runs); GitHub silently registered the workflows with no triggers. Fixed by replacing the here-strings with string-array `-join`. `gh workflow list` now shows both with their proper friendly names.
- **Release-notes section lookup fixed**: both release workflows now read the `[<version>]` CHANGELOG section (matching the established convention of rolling `[Unreleased]` &rarr; `[<version>]` in the version-bump commit) instead of `[Unreleased]`.
- Setup gate: `PSGALLERY_API_KEY` repository secret was added this session as a one-time configuration.

**Bug fixes shipped in 4.7.2:**

- **`Set-VisioUserDefinedCell`** &mdash; `-Value` and `-Prompt` parameters now encode their string arguments as Visio formulas via the new `SetString` typed setter from #144. The cmdlet was assigning raw strings to `.Value` / `.Prompt` since commit `eb50bff1` (years old; latent because anyone hitting the path could work around with pre-quoted values like `-Value '"foo"'`). Pre-#144 the cmdlet threw `COMException: #NAME?`; #144's detect-and-rethrow in 4.7.0 wrapped that as a friendly `ArgumentException` but didn't fix the cmdlet's underlying behavior. New PS regression test: `VisioPS_SetVisioUserDefinedCell_EncodesValueAndPrompt`. Not a 4.7.0 regression; long-standing bug discovered while doing #144 follow-up.
- **`CmdletsToExport='*'` &rarr; explicit list of 64 cmdlets** in [`Visio.psd1`](VisioAutomation_2010/VisioPowerShell/Visio.psd1). Resolves one of the three publish-time warnings emitted by `Publish-Module`. Module-load is also marginally faster (no wildcard reflection at import time).

**Still-pending publish-time warning** (futures item, customer-impact analysis pre-derived): the `ModuleToProcess` &rarr; `RootModule` rename + `PowerShellVersion = '2.0'` &rarr; `'5.1'` bump. **Effective customer impact: zero.** PS 2.0 was removed from Windows 11; PSGallery requires PS 5.1+; the binary is compiled against PS 3.0 reference assemblies and won't load on PS 2.0 regardless of what the manifest claims. Bumping the manifest aligns the declaration with what the module actually requires; LTSB 2016 (the ongoing compat constraint until 2026-10-13) ships with PS 5.1, so unaffected. Detail in [`docs/futures/releases.md`](docs/futures/releases.md).

**Several new futures items captured this session** for forward-looking work:

- **Move to C# 14 / .NET 10** ([`build-and-code.md`](docs/futures/build-and-code.md)) &mdash; concrete landing-point plan replacing the older "Decide whether to move to .NET 6/8". Implications analyzed for VisioPowerShell (the binary module won't load in Windows PowerShell 5.1 on .NET 10; multi-target with `net48` + `net10.0-windows` is the recommended path). C# 14 features worth using (extension members are the headline payoff, given the codebase's extension-method surface). Sequenced after VS 2026 + 4.7.2 TFM bump (gated on LTSB 2016 sunset 2026-10-13).
- **Borrow ideas from VisioBot3000 for VisioPS ergonomics** ([`build-and-code.md`](docs/futures/build-and-code.md)) &mdash; nickname registry, dynamic function generation per registered shape, block-style nested syntax for containers, relative positioning cursor. Four-phase adoption path; the dynamic-function DSL is Phase 4 (deferred until earlier phases prove the demand).
- **Borrow ideas from PSVA for VisioPS bulk-operation cmdlets** ([`build-and-code.md`](docs/futures/build-and-code.md)) &mdash; `Set-VisioShapeDistribution`, pipeline-friendly bulk connectors, side-and-alignment shape decoration, layer cmdlets. Half-day audit pass needed first to map against existing VisioPS surface.
- **Evaluate NetOffice / NetOfficeFw as a Visio PIA replacement** ([`build-and-code.md`](docs/futures/build-and-code.md)) &mdash; either adopt directly or pattern-mine the COM-cleanup `IDisposable` shape. NetOfficeFw modern-.NET support status is a load-bearing input to *Move to C# 14 / .NET 10*.
- **Address `Visio.psd1` deprecation warnings** ([`releases.md`](docs/futures/releases.md)) &mdash; partial; `CmdletsToExport` half landed this session, `ModuleToProcess`/`PowerShellVersion` half pending with customer-impact analysis pre-derived.

**Repo state:** master at HEAD (whatever the CLAUDE.md-update commit is). Working tree clean. PSGallery: `Visio` 4.7.2 live. NuGet: `VisioAutomation2010` 3.0.0 live (shipped this session). `[Unreleased]` in [`NuGet/CHANGELOG.md`](NuGet/CHANGELOG.md) is now empty &mdash; ready for the next cycle. Both gitbook repos in sync with origin.

**Test infrastructure:** all four projects green &mdash; VTest 101/101, VTest.Models 59/59, VTest.Scripting 34/34, VTest.PowerShell 20/20. Total 214. (Up from 177 baseline.) New this session: `VisioPS_Manifest_CmdletsToExport_HasNoMissingEntries` and `VisioPS_Manifest_CmdletsToExport_HasNoExtraEntries` in [`VisioPS_Manifest_Tests.cs`](VisioAutomation_2010/VTest.PowerShell/VisioPS_Manifest_Tests.cs) &mdash; pure metadata drift check between [`Visio.psd1`](VisioAutomation_2010/VisioPowerShell/Visio.psd1)'s `CmdletsToExport` and the `[Cmdlet]`-attributed types in `VisioPS.dll`. Mirrors the same check `publish-psmodule.yml` runs at publish time, but at unit-test time so drift is caught on every test run instead of only at release. No Visio runtime needed (reflection + `PSParser.Tokenize`).

**Dev-environment note:** Smart App Control on Windows 11 can transiently block freshly-built test DLLs (`Application Control policy has blocked this file`, HRESULT 0x800711C7, Code Integrity policy ID `{0283ac0f-fff1-49ae-ada1-8a933130cad6}`). The block clears once SAC's cloud reputation cache catches up (30-60s); VS 2022 Test Explorer hits the same mechanic. Durable fix: turn Smart App Control off via Settings &rarr; Privacy & security &rarr; Windows Security &rarr; App & browser control &rarr; Smart App Control settings &rarr; Off. **One-way operation** &mdash; re-enabling requires reinstalling Windows. The CI runners (GitHub-hosted `windows-latest`) are not subject to this.

**Prior session context** (still recent, still relevant):

- **Issue [#144](https://github.com/saveenr/VisioAutomation/issues/144) closed.** Custom property typed setters (`SetString` / `SetNumber` / `SetBool` / `SetDate` / `SetFormula` on `CustomPropertyCells`; `SetString` / `SetFormula` on `UserDefinedCellCells`); `Value` &rarr; `Formula` rename with `[Obsolete]` shim; `CustomPropertyHelper.Set` and `UserDefinedCellHelper.Set` wrap `COMException: #NAME?` as `ArgumentException` with a friendly diagnostic. Engineering reference at [`docs/internal/custom-property-encoding.md`](docs/internal/custom-property-encoding.md). User-facing matrix on the [.NET](https://saveenr.gitbook.io/visioautomation/custom-properties) and [PS](https://saveenr.gitbook.io/visiopowershell/automatic-diagrams/drawing-directed-graphs#adding-custom-properties-to-nodes) gitbooks. [#117](https://github.com/saveenr/VisioAutomation/issues/117) reporter pinged with the one-line fix; [#145](https://github.com/saveenr/VisioAutomation/issues/145) ready to close.
- **Phase 3 SDK migration done.** All 11 csprojs SDK-style + PackageReference + Central Package Management. Detail in [`docs/COMPLETED.md`](docs/COMPLETED.md#phase-3--modernization-in-progress).
- **Open-issues backlog pass** closed nine issues (#105 / #128 / #129 / #130 / #138 / #139 / #140 / #141 / #142 / #143) and shipped the directed-graph XML attribute work.

**Next session priorities:**

1. **NetOffice spike** (~1 day) &mdash; investigate whether `NetOfficeFw.Visio` covers the codebase's Visio-COM surface and ships modern-.NET TFMs. Output is a go/no-go memo. High-information-value for the upcoming `Move to C# 14 / .NET 10` work. Independent and bounded. See [`docs/futures/build-and-code.md`](docs/futures/build-and-code.md).
2. **Address remaining `Visio.psd1` deprecation warnings** &mdash; the `ModuleToProcess` &rarr; `RootModule` rename + `PowerShellVersion '2.0'` &rarr; `'5.1'` bump. Customer impact already analyzed (zero); see [`docs/futures/releases.md`](docs/futures/releases.md). One-line change times two; could ride along with any future PS module patch. **Note:** when this is picked up, also re-evaluate the version-policy decision (currently "stay divergent" &mdash; the PS-5.1 manifest bump is the agreed forcing function for revisiting convergence; see [`docs/futures/releases.md`](docs/futures/releases.md#reconcile-version-numbers-across-artifacts-decision-stay-divergent-for-now-revisit-at-ps-51-manifest-bump)).
3. **Small follow-ups deferred from the migration** &mdash; the `<frameworkAssembly assemblyName="Microsoft.Office.Interop.Visio" />` line in [`NuGet/VisioAutomation2010.nuspec`](NuGet/VisioAutomation2010.nuspec) is redundant with the bundled-PIA-via-NuGet behavior. Trim as a small standalone commit when convenient.
4. **NuGet support case to declassify the `saveenr` account** *(low priority, optional)* &mdash; the `SevenPens` workaround works fine, but if you ever want `saveenr` to be the canonical publisher again (e.g., for symmetry with the gitbooks / GitHub repo identity), opening a NuGet support case is the path. See the gate-quirk write-up in [`docs/futures/releases.md`](docs/futures/releases.md#microsoft-package-compliance-gate-on-the-saveenr-nugetorg-account-operational-quirk-discovered-2026-05-07-during-the-300-publish).
5. **(Post-2026-10-13) TFM bump** &mdash; only after Windows 10 LTSB 2016 leaves Extended Support. See [enterprise compat memory](../../.claude/projects/C--Users-savee-Documents-GitHub-VisioAutomation/memory/enterprise_compat_ltsb2016.md). Unblocks VS 2026 move and ultimately the *Move to C# 14 / .NET 10* item.

## Other docs in this repo

- [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md) — projects, dependencies, central concepts
- [docs/BUILDING.md](docs/BUILDING.md) — full build / test / install reference (incl. dev-pack install commands)
- [docs/TESTING.md](docs/TESTING.md) — test-suite design and conventions (shared `Framework.VTest` base, per-testhost Visio singleton, `[AssemblyCleanup]` orphan-prevention, MSTEST0030 enforcement). Per-project READMEs sit next to each test csproj.
- [docs/GLOSSARY.md](docs/GLOSSARY.md) — Visio and codebase terminology
- [docs/ROADMAP.md](docs/ROADMAP.md) — staged plan (Phase 1 / 2 / 3); read first for orientation
- [docs/MILESTONES.md](docs/MILESTONES.md) — themed milestones with proposed target windows. Sits between ROADMAP.md (phases) and FUTURES.md (backlog index)
- [docs/FUTURES.md](docs/FUTURES.md) — index of backlog items, split by topic into [`docs/futures/*.md`](docs/futures/)
- [docs/OVERVIEW.md](docs/OVERVIEW.md) — entry point, links to the above
