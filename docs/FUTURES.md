# Futures — 2026 Refresh Backlog

A running list of cleanup, modernization, and improvement items for the VisioAutomation solution. Items are grouped by theme. Each entry includes a one-line **What**, a **Why** (cost of leaving it), and a rough **Effort** (S / M / L). This is a *backlog* — items are not committed to or scheduled until pulled out into actual work.

---

## Roadmap (staged plan)

The 2026 refresh runs in three phases. Each backlog item below is tagged with its phase.

### Phase 1 — VS 2022 cleanup *(done; merged to master 2026-05-03)*
Stayed on Visual Studio 2022 and the current TFMs (.NET Framework 4.5.2 for shipping libs, 4.7.2 for tests). Code + docs improvements only, no new features. The phase culminated in the **Visio PowerShell 4.6.1** release on 2026-05-03 (tag `VisioPS_4.6.1`).

Phase 1 items completed:
- ✅ *Revise user-facing documentation for accuracy* — full audit and rewrite of [VisioPowerShellDocs](https://saveenr.gitbook.io/visiopowershell) and the .NET-side gitbook docs. Standardized every cmdlet page on a Syntax + Parameters + Examples + See-also layout. Reader-facing summary at [`documentation-changes.md`](https://saveenr.gitbook.io/visiopowershell/documentation-changes).
- ✅ *Cmdlet bug fixes shipped in 4.6.1* — `Lock-VisioShape` / `Unlock-VisioShape` switches now actually bind; `Export-VisioShape` file-exists check no longer inverted; `New-VisioShape` polyline / Bezier minimum-point validation actually throws.
- ✅ *Manual release machinery* — [`Publish-VisioPSToGallery.ps1`](../VisioAutomation_2010/VisioPowerShell/Publish-VisioPSToGallery.ps1) wraps the staging / publish / tag / push flow with TLS 1.2 forcing, `-ErrorAction Stop`, and post-publish gallery verification. Documented in [VisioPowerShellDocs/developer-info/publishing-to-powershell-gallery.md](https://saveenr.gitbook.io/visiopowershell/developer-info/publishing-to-powershell-gallery).
- ✅ *Fix the misnamed PowerShell loader script* — rewrote it to actually `Save-Module` from the PS Gallery
- ✅ *Add a `CLAUDE.md` at the repo root* — added with staged-plan, build commands, conventions, doc pointers
- ✅ *Update MSTest off the beta* — upgraded `MSTest.TestFramework` and `MSTest.TestAdapter` from `2.0.0-beta2` to `4.2.2`; bumped `VTest` TFM 4.5 → 4.7.2 to satisfy MSTest 4.x's floor
- ✅ *Add a per-project `README.md` for the larger projects* — `VisioAutomation/`, `VisioAutomation.Models/`, `VisioScripting/`, `VisioPowerShell/` (already had one)
- ✅ *Add a `CONTRIBUTING.md`* — covers branch, setup pointer, tests-need-Visio rule, code style, commits, changelog discipline, per-phase scope
- ✅ *Expand the root `readme.md`* — rewrote with pitch, install table, C# + PowerShell quick-start, doc links, license
- ✅ *Audit `Internal/` for dead code* — deleted orphaned `TempHelper.cs` + removed dead `InternalsVisibleTo("TestVisioAutomation")` attribute; spawned a follow-up item for misc warts found during the audit
- ✅ *Misc cleanups discovered during the Internal/ audit* (mostly) — moved misplaced `InternalsVisibleTo` attributes to `AssemblyInfo.cs`, deleted two orphaned VTest files, removed auto-generated `.sln.metaproj` from version control. `LinqExtensions` visibility-vs-folder mismatch deferred to Phase 3 as a breaking-namespace-change risk.
- ✅ *Add CI* (build-only) — `.github/workflows/build.yml` builds the solution on push/PR for `master`, pinned to VS 2022 MSBuild, NuGet packages cached. Test runs in CI deferred to Phase 3 (needs self-hosted runner with Visio).

### Phase 2 — Cut the final release
Tag and publish a final release of VisioAutomation (NuGet) with the refreshed docs. The PowerShell-module half of this phase shipped early as **Visio PowerShell 4.6.1** on 2026-05-03; only the NuGet release remains.

Phase 2 prerequisites (must be settled before the NuGet release ships):
- *Reconcile version numbers across artifacts* — needs a deeper conversation before a decision; **currently deferred**, do not implement until discussed. The PS module is now at `4.6.1`; the NuGet is at `2.6.0`.
- *Investigate flakiness from leftover Visio processes* — relevant to the release-verification flow.

### Phase 3 — Modernization
- *Move development to Visual Studio 2026*
- *Consolidate target frameworks* — step 2 (4.5 → 4.7.2)
- *Consider migrating off Visio 2010 PIA*
- *Decide whether to move to .NET 6/8 (out of .NET Framework)*
- *Migrate from `packages.config` to `PackageReference`*
- *Modernize SDK-style csproj*
- *Automate releases via GitHub CI — NuGet + PowerShell Gallery*
- *Decide where docs live long-term*

---

## Build & tooling

### Consolidate target frameworks
- **Status:** Step 1 done. All shipping libraries (`VisioAutomation`, `VisioAutomation.Models`, `VisioScripting`, `VisioPowerShell`) and both sample projects (`VSamples`, `VSamples.Docs`) are now on **.NET Framework 4.5.2** (originally bumped from 4.0 → 4.5 in commit `2fd6b466`, then 4.5 → 4.5.2 to satisfy the available Developer Pack — see BUILDING.md). Test projects on **.NET Framework 4.7.2** (VTest moved there as part of the MSTest 4.x upgrade; the others were already there).
- **Step 2 (remaining):** Bump the shipping fleet again to clear the **VS 2026** floor (Framework 4.6.2 minimum). Recommended landing point: **4.7.2** — same TFM the test projects already use, so the whole solution converges on one number. **Side benefit of step 2:** the .NET Framework 4.5.2 Developer Pack will no longer be required on dev machines or CI runners (currently it is — see [BUILDING.md](BUILDING.md) prereqs). The v4.7.2 reference assemblies ship in-box on every supported Windows. See *Move development to Visual Studio 2026* below.
- **Why:** Mixed TFMs cause subtle binary-compatibility surprises (a test project on a higher TFM can use APIs the library under test cannot). Step 1 eliminated the production 4.0/4.5 split; step 2 will eliminate the 4.5/4.7.2 split between shipping libs and tests, and let us drop the v4.5 reference-assemblies NuGet workaround.
- **Effort:** S (already partially done).

### Migrate from `packages.config` to `PackageReference`
- **What:** Every csproj still uses the old `packages.config` NuGet model.
- **Why:** `PackageReference` is transitive, lockable, and the only model supported by `dotnet` CLI / SDK-style projects. Required before any modernization beyond Framework.
- **Effort:** S–M

### Update MSTest off the beta ✅ done
- **What:** All test projects pinned `MSTest.TestFramework` and `MSTest.TestAdapter` to `2.0.0-beta2`.
- **Resolution:** Upgraded both packages to **4.2.2** (latest stable) across all four test projects. Required bumping VTest from .NET Framework 4.5 → 4.7.2 (MSTest 4.x's floor is 4.6.2); the other test projects were already on 4.7.2. Solution builds cleanly. Test code did not need any changes — MSTest's `[TestMethod]`/`Assert.*` API surface is stable across the version jump. Tests not actually run (need a live Visio).

### Add CI ✅ done (build-only)
- **Resolution:** [`.github/workflows/build.yml`](../.github/workflows/build.yml) added. Builds the solution in Debug on every push to `master` / `2026_Refresh` and on every PR. Pinned to VS 2022's MSBuild (`vs-version: '17.0'`) since VS 2026's MSBuild can't resolve the .NET Framework 4.5.2 reference assemblies the shipping libs need. The workflow also installs the .NET Framework 4.5.2 Developer Pack via chocolatey before building (those reference assemblies aren't on the runner image). NuGet packages are cached keyed on the hash of all `packages.config` files. Build status surfaces as a badge in the root README.
- **Still to do (Phase 3):** Run the tests in CI. This needs a self-hosted Windows runner with Microsoft Visio installed. Track alongside *Automate releases via GitHub CI* in Phase 3.

### Modernize SDK-style csproj
- **What:** Convert the legacy csproj format (long `<Compile Include="..." />` lists, packages.config) to SDK-style csproj.
- **Why:** Smaller files, no need to enumerate every source file, easier diffs, prerequisite for any later .NET migration.
- **Effort:** M (depends on PackageReference being done first).

---

## Code & architecture

### Consider migrating off Visio 2010 PIA
- **What:** All projects reference `Microsoft.Office.Interop.Visio` v14 (Visio 2010 PIA). Visio is now on a much newer version (16.x, with Visio for Microsoft 365).
- **Why:** The 2010 PIA still works at runtime against newer Visio versions, so this isn't urgent. But APIs added since 2010 are inaccessible without rebinding to a newer interop assembly. Decide whether to stay on 2010 (max compatibility) or move forward (access to newer features).
- **Effort:** M — touches every project; needs a compatibility decision.

### Decide whether to move to .NET 6/8 (out of .NET Framework)
- **What:** Whole solution is .NET Framework. Modern .NET supports COM interop on Windows.
- **Why:** Long-term viability — .NET Framework only gets security updates. But COM interop on modern .NET has its own quirks, and the PowerShell module bridge (Windows PowerShell 5.1 vs PowerShell 7) becomes a bigger decision.
- **Effort:** L — major undertaking; do PackageReference + SDK-style first.

### Audit `Internal/` for dead code ✅ done
- **Resolution:** Audited `VisioAutomation/Internal/` end-to-end. Two clear-cut wins removed: deleted [`TempHelper.cs`](../VisioAutomation_2010/VisioAutomation/Internal/) (orphaned — duplicate of `ShapesheetHelpers` with snake_case names; not even listed in the csproj, so already wasn't being compiled) and removed the dead `[assembly: InternalsVisibleTo("TestVisioAutomation")]` attribute from `AssemblyInfo.cs` (no `TestVisioAutomation` assembly exists; current test projects are `VTest` / `VTest.Scripting`, granted access elsewhere). All other Internal/ types are actively referenced. Build verified clean.

### Misc cleanups discovered during the Internal/ audit ✅ done (mostly)
- **Resolution:** Three of the four findings cleaned up:
  - ✅ Moved `[InternalsVisibleTo("VTest")]` and `[InternalsVisibleTo("VTest.Scripting")]` from `Internal/ArraySegmentEnumerator.cs` to `Properties/AssemblyInfo.cs` where they belong.
  - ✅ Deleted `Vtest/FormatStringParserTest.cs` and `VTest/Core/Extensions/AsEnumerableTest.cs` — turned out to be fully orphaned (not in csproj, contained syntax errors and references to long-gone types like `Isotope.Text.FormatStringParser` and a `VisioAutomationTest` base class). Pure dead code.
  - ✅ Removed `VisioAutomation2010.sln.metaproj` from version control and added `*.sln.metaproj` / `*.sln.metaproj.tmp` to `.gitignore` so it gets regenerated on demand instead of accumulating stale paths.
- **Deferred to Phase 3:**
  - **`LinqExtensions` is `public` despite living in `Internal/Extensions/`.** It's actually consumed across the assembly boundary by `VisioAutomation.Models` (`ShapeList` uses its single `NotOfType<T>` method). The `public` visibility is therefore correct, but the folder name is misleading. Fix is to either move it out of `Internal/` (a namespace change, technically a breaking API change) or rename the folder.

---

## Tests

### Tests require a live Visio
- **What:** Every test project spins up a real Visio process via COM. There is no mock/fake layer.
- **Why (consider):** This is intentional — the library's whole job is to drive Visio, and mocking COM gives false confidence. But the lack of any non-Visio test surface means there's no quick `dotnet test` that runs anywhere. *Not necessarily a problem*, just worth a deliberate decision before adding CI.
- **Effort:** N/A — design decision, not a task.

### Investigate flakiness from leftover Visio processes *(Phase 2 prereq)*
- **What:** Aborted test runs can leave Visio processes that lock files and break the next run.
- **Why:** Add a test-host shutdown hook or pre-run cleanup so re-runs are deterministic. Important for the release-verification flow in Phase 2 — re-running the test suite should be reliably idempotent before we ship.
- **Effort:** S.

---

## Packaging & versioning

### Reconcile version numbers across artifacts *(Phase 2 prereq — deferred, needs discussion)*
- **What:** The NuGet [`VisioAutomation2010.nuspec`](../NuGet/VisioAutomation2010.nuspec) is at `2.6.0`; the PowerShell [`Visio.psd1`](../VisioAutomation_2010/VisioPowerShell/Visio.psd1) is at `4.6.0`; csproj `AssemblyVersion`s are independent again.
- **Why:** Hard to tell at a glance which library version corresponds to which module version. Same code (the PS module bundles the NuGet's DLLs) shipping under two different version numbers makes bug reports ambiguous and release coordination loose.
- **Status (2026-05-03):** Held for further discussion. Three options on the table:
  - **A — Converge:** both artifacts ship at the same number going forward; pick a number for the Phase 2 release.
  - **B — Document the divergence policy:** keep them independent, write down the rule.
  - **C — Single technical source of truth:** `Directory.Build.props` + token substitution into nuspec/psd1. Better suited to Phase 3 once csprojs are SDK-style.
- **Forcing function:** must be answered before the Phase 2 release ships, since release time is when the version strings get bumped.
- **Effort:** S for the policy decision and doc updates; M if Option C is chosen.

### Fix the misnamed PowerShell loader script ✅ done
- **What:** `DownloadFromPowerShellGallery.ps1` did not download from the PowerShell Gallery — it `Import-Module`d the local `bin\Debug` build.
- **Resolution:** Rewrote the script to `Save-Module Visio` from PSGallery into a local `DownloadedModule/` subfolder (gitignored) and `Import-Module` that. Now serves as a one-shot release-verification helper. Renamed the file to [`LoadFromGallery.ps1`](../VisioAutomation_2010/VisioPowerShell/LoadFromGallery.ps1) for parallelism with the other loader scripts; also renamed [`LoadFromBinDebug_ISE.ps1`](../VisioAutomation_2010/VisioPowerShell/LoadFromBinDebug.ISE.ps1) → `LoadFromBinDebug.ISE.ps1` (dot-suffix is more idiomatic than underscore).

### Publish the PowerShell module to the PowerShell Gallery
- **What:** The module is currently distributed only by manual install (`InstallForCurrentUser.ps1`).
- **Why:** Gallery publication makes `Install-Module Visio` work for users. Requires deciding on the publication identity, signing, and a release process.
- **Effort:** M — operational rather than coding work.

### Switch module-release builds from Debug to Release
- **What:** The release-prep script [`InstallForCurrentUser.ps1`](../VisioAutomation_2010/VisioPowerShell/InstallForCurrentUser.ps1) hardcodes `$release = "Debug"` (line 69). The 4.6.1 release was published from the Debug build to keep the workflow unchanged, but for future releases we should ship the Release build — smaller binaries, no `DEBUG` symbols, no JIT debug overhead.
- **Why:** Shipping Debug builds to consumers is sloppy hygiene. Should be Release for any artifact that goes to a public feed (PSGallery, NuGet).
- **How:** Either flip the constant in `InstallForCurrentUser.ps1` (and document in the script comment that release-mode is now used for actual releases), or split the script into `InstallForCurrentUser.ps1` (Debug, dev convenience) and a separate `Stage-ReleaseBuild.ps1` (Release, used by `Publish-VisioPSToGallery.ps1`).
- **Effort:** S.

### Publish the NuGet package to nuget.org
- **What:** Same question for the NuGet package as for the PS module.
- **Effort:** S–M.

---

## Documentation

### Add a `CLAUDE.md` at the repo root ✅ done
- **What:** Project-specific instructions for future Claude Code sessions: build commands, test rules (need Visio installed), where the public API lives, the `2026_Refresh` branch convention.
- **Resolution:** [`CLAUDE.md`](../CLAUDE.md) added. Covers the staged-plan summary, verified build commands, test prerequisites, the per-commit changelog convention, the PS loader-script naming convention, tooling notes (shell choice, MSBuild path), and pointers to the rest of the docs.

### Add a `CONTRIBUTING.md` ✅ done
- **Resolution:** [`CONTRIBUTING.md`](../CONTRIBUTING.md) added at the repo root. Covers the active branch convention, setup pointer to BUILDING.md, the live-Visio test rule, code style guidance (don't reformat, no new files unless needed, default to no comments), commit message format, the per-commit changelog convention, and the per-phase scope rules from FUTURES so contributors don't accidentally violate them.

### Expand the root `readme.md` ✅ done
- **Resolution:** Rewrote the three-line readme into a proper landing page: NuGet/PSGallery/license badges, the elevator pitch, an artifact table with install commands, two quick-start examples (C# + PowerShell), links to user guides, in-repo developer docs, release notes, contributing, and license.

### Decide where docs live long-term
- **What:** User docs are in a separate repo on gitbook ([`VisioAutomation_GitBook_Docs`](https://github.com/saveenr/VisioAutomation_GitBook_Docs)); developer docs are now in `docs/` here.
- **Why:** Two-repo doc setups drift. Either keep them split with a clear policy (which doc lives where) or consolidate. No urgent action needed — just call out the policy in `OVERVIEW.md` once decided.
- **Effort:** S (policy) — or M (consolidation).

### Expand .NET-side doc coverage — Tier 3 (`VisioAutomation.Models`)
- **What:** The 2026 audit on [`VisioAutomation_GitBook_Docs`](https://github.com/saveenr/VisioAutomation_GitBook_Docs) reviewed every existing page for accuracy and added 15 new pages over three tiers. Tier 3 is the only group still pending.
- **Tier 1 — common helpers** *(done)*: Hyperlinks, Lock cells, Control handles, Connection points, Connectors.
- **Tier 2 — structural cell-helper pages** *(done)*: Shape format / layout / xform cells, Page cells, Text formatting, Geometry, Application.
- **Tier 4 — smaller / niche public surface** *(done)*: Analyzers, Visio error log (LoggingHelper / XmlErrorLog), UndoScope, Exception types, plus a full rewrite of `extension-methods.md` covering all 16 `Extensions/` method classes (LINQ bridges, drawing primitives, drop, ShapeSheet I/O, geometry / coordinates, one-offs).
- **Why Tier 3 still:** It's the most useful unwritten chunk &mdash; `VisioAutomation.Models` covers the high-level "build a diagram declaratively / render it" flow that powers the `Out-VisioApplication` cmdlet on the PS side. Library users currently have to read the source to discover `OrgChartDocument`, `DirectedGraphDocument`, the layout-style classes, the DOM model, etc.
- **Tier 3 page list (~6–8 pages):**
  - **DOM document model** — `Document`, `Page`, `MasterRef`, `Connector`, `Line`, `Oval`, `BezierCurve`, `PolyLine`, `Hyperlink`, the `Node`/`NodeList` containment pattern. The declarative way to build a Visio document.
  - **Layouts** — `LayoutStyleBase` and its subclasses (`FlowchartLayoutStyle`, `RadialLayoutStyle`, `CompactTreeLayoutStyle`, `HierarchyLayoutStyle`, `CircularLayoutStyle`, `OrganizationalChartLayoutStyle`).
  - **OrgChart** — `OrgChartDocument`, `OrgChartStyling`, `OrgChartLayoutOptions`. The model side of the existing `Out-VisioApplication -OrgChart` flow on the PowerShell side.
  - **DirectedGraph** — `DirectedGraphDocument` and node/edge types. The richer of the two graph models.
  - **DataTable** — `DataTableModel` for tabular layouts.
  - **XmlModel** — generic XML-backed renderer.
  - **Forms** — `FormDocument`, `FormPage`, `InteractiveRenderer`, `TextBlock` (the lightweight form-builder). Probably worth one page.
- **Effort:** M (6–8 pages).
- **How to apply:** Same pattern as Tiers 1 / 2 / 4: one paragraph of conceptual framing, a field/method table when the surface is bigger than two methods, code examples for the common operations. Each new page goes into [SUMMARY.md](https://github.com/saveenr/VisioAutomation_GitBook_Docs/blob/main/SUMMARY.md) and gets a one-line entry in [`documentation-changes.md`](https://github.com/saveenr/VisioAutomation_GitBook_Docs/blob/main/documentation-changes.md) under "Pages added".

### Keep CHANGELOGs current as Phase 1 work lands
- **What:** Two changelogs were added in [Keep a Changelog](https://keepachangelog.com/en/1.1.0/) format: [`NuGet/CHANGELOG.md`](../NuGet/CHANGELOG.md) for the `VisioAutomation2010` NuGet, and [`VisioAutomation_2010/VisioPowerShell/CHANGELOG.md`](../VisioAutomation_2010/VisioPowerShell/CHANGELOG.md) for the `Visio` PowerShell module. Each has an `[Unreleased]` section that should accumulate consumer-visible changes until the Phase 2 release cuts a real version.
- **Why:** The whole point of cutting a final release in Phase 2 is to give consumers a clean, well-documented checkpoint. If Unreleased sections drift behind reality during Phase 1, the release notes will be wrong.
- **How to apply:** When a Phase 1 commit changes anything a consumer of the NuGet or PS module would notice (public API, parameter behavior, supported runtime, dependencies), add an entry to the corresponding CHANGELOG's `[Unreleased]` in the same commit. Pure internal/build/docs changes don't need entries.
- **Effort:** ~zero per change, if done in the same commit.

### Add a per-project `README.md` for the larger projects ✅ done
- **Resolution:** READMEs added for all four larger projects: [`VisioAutomation/`](../VisioAutomation_2010/VisioAutomation/README.md), [`VisioAutomation.Models/`](../VisioAutomation_2010/VisioAutomation.Models/README.md), [`VisioScripting/`](../VisioAutomation_2010/VisioScripting/README.md), [`VisioPowerShell/`](../VisioAutomation_2010/VisioPowerShell/README.md). Each covers folder layout, key types where relevant, and pointers to ARCHITECTURE / GLOSSARY / BUILDING / CHANGELOG.

---

## Items added during discussion

### Move development to Visual Studio 2026
- **What:** Bump the solution from VS 2022 (`VisualStudioVersion = 17.0` in the .sln) to VS 2026. Stay on .NET Framework — do not migrate to modern .NET (Core).
- **Constraint discovered during research:** VS 2026 supports .NET Framework targets **4.6.2, 4.7, 4.7.1, 4.7.2, 4.8, 4.8.1** only. Framework 4.0 / 4.5 / 4.5.x / 4.6 / 4.6.1 are **not** supported targets in VS 2026. Source: [Visual Studio 2026 Compatibility](https://learn.microsoft.com/en-us/visualstudio/releases/2026/compatibility).
- **Implication:** the shipping fleet (currently on 4.5.2 after step 1 of *Consolidate target frameworks*) must bump again before VS 2026 can build it. Recommended landing point: **4.7.2** — clears the VS 2026 floor *and* converges with the existing test-project TFM in one move.
- **VisioPowerShell older-PowerShell support is preserved** by this bump: the older-PS floor is set by the `System.Management.Automation` v3 reference and the `ModuleToProcess`/`PowerShellVersion = 2.0` choices in [Visio.psd1](../VisioAutomation_2010/VisioPowerShell/Visio.psd1), not by the .NET Framework TFM. Bumping 4.5 → 4.7.2 doesn't change that.
- **Cross-refs:** Drives step 2 of *Consolidate target frameworks*. Supersedes *Decide whether to move to .NET 6/8* for now (defer that decision).
- **Effort:** S — bump TFMs, bump `VisualStudioVersion` in the .sln, open in VS 2026, full rebuild.

### Automate releases via GitHub CI *(immediate next item)*
- **What:** Replace the current manual release process with a GitHub Actions workflow (or set of workflows) that handles **three deliverables** end-to-end:
  1. **PSGallery publish** of the `Visio` PowerShell module.
  2. **nuget.org publish** of the `VisioAutomation2010` NuGet package.
  3. **GitHub Release** with the built binaries (DLLs / `.zip` / the `.nupkg`) attached as downloadable artifacts.
- **Why:** Manual releases are error-prone and infrequent. The 4.6.1 release surfaced several PS-5.1 / PowerShellGet gotchas that are now baked into [`Publish-VisioPSToGallery.ps1`](../VisioAutomation_2010/VisioPowerShell/Publish-VisioPSToGallery.ps1) and the [Publishing doc](https://saveenr.gitbook.io/visiopowershell/developer-info/publishing-to-powershell-gallery); automating those steps ensures future releases inherit the workarounds. GitHub Releases also give consumers a stable URL to download a specific version's binaries even if PSGallery / nuget.org are slow to update.

#### References for the workflow content

- **PSGallery publish** &mdash; [`Publish-VisioPSToGallery.ps1`](../VisioAutomation_2010/VisioPowerShell/Publish-VisioPSToGallery.ps1) is the canonical battle-tested flow (TLS 1.2, `-Path` not `-Name`, `-ErrorAction Stop`, post-publish verification via `Find-Module`, then tag). It's callable from the workflow as-is; reads the API key from `$env:PSGalleryApiKey` or `-ApiKey`.
- **NuGet publish** &mdash; the package metadata is in [`NuGet/VisioAutomation2010.nuspec`](../NuGet/VisioAutomation2010.nuspec) (currently `2.6.0`). No equivalent battle-tested script exists; the workflow needs a `nuget pack` + `nuget push` step or the equivalent `dotnet nuget push`. NuGet's gallery and `Publish-Module` don't share infrastructure &mdash; expect different gotchas.
- **GitHub Release** &mdash; the [`softprops/action-gh-release@v2`](https://github.com/softprops/action-gh-release) action handles upload-on-tag-push idiomatically. Alternative: `gh release create`.
- **Existing CI infrastructure** &mdash; [`.github/workflows/build.yml`](../.github/workflows/build.yml) is the build-only workflow. Pinned to VS 2022 MSBuild via `microsoft/setup-msbuild@v2`; installs `netfx-4.5.2-devpack` via chocolatey before building. The release workflow should copy this setup verbatim.

#### Decisions to make first

- **One workflow or three?** Cleanest: a single `release.yml` triggered on `workflow_dispatch` (or tag) with conditional steps based on which deliverable to ship. Feature-flagged by inputs makes the first-cut testing easier.
- **Trigger.** Three reasonable choices:
  - `workflow_dispatch` with version + flags as inputs (manual, fully controlled, recommended for first cut).
  - Tag-push: `VisioPS_*` &rarr; PSGallery + GitHub Release for the module bundle; `VisioAutomation_*` &rarr; NuGet + GitHub Release for the library. Two separate workflow files keyed on tag pattern.
  - GitHub Release creation as the trigger (`on: release: types: [published]`). Less appealing because the manual `Publish-VisioPSToGallery.ps1` script already creates the tag at the end of a successful publish; reproducing the same ordering means the workflow creates the GitHub Release at the end too.
- **Tag-then-publish vs. publish-then-tag.** The 4.6.1 manual flow tagged **after** verifying the publish landed. Reproducing that ordering in CI pushes toward `workflow_dispatch` (publish, then tag from inside the workflow) rather than tag-push. Note: a subsequent GitHub Release creation step would then attach to that tag.
- **What artifacts go into the GitHub Release?** Candidates: the staged module folder zipped (the same content that's published to PSGallery), the `.nupkg` from the NuGet publish, possibly a separate "binaries-only" zip of the DLLs for users who don't want either package manager. Keep it small to start; one zip with the module is sufficient as a v1.
- **Build configuration.** Phase 1 shipped 4.6.1 from a Debug build (`InstallForCurrentUser.ps1` hardcodes `Debug`). Future releases should switch to Release; tracked in [the *Switch module-release builds from Debug to Release* item](#switch-module-release-builds-from-debug-to-release). The CI workflow either flips the constant or stages the release config separately.
- **Signing.** Authenticode signing of the bundled DLLs is open. Required by neither PSGallery nor nuget.org but would silence the "publisher unknown" warning. Defer until the workflow is otherwise stable.
- **Version policy.** Module is at `4.6.1`; NuGet is at `2.6.0`. Until [*Reconcile version numbers across artifacts*](#reconcile-version-numbers-across-artifacts-phase-2-prereq--deferred-needs-discussion) is settled, the workflow has to handle two different version sources (read PS module version from `Visio.psd1`, NuGet version from `VisioAutomation2010.nuspec`). That's fine; just be explicit about it.

#### Subtasks

- **Confirm credentials and ownership:**
  - PSGallery: `Visio` package, key stored as GitHub secret (suggested name: `PSGALLERY_API_KEY`).
  - nuget.org: `VisioAutomation2010` package &mdash; confirm ownership and add the secret (suggested name: `NUGET_API_KEY`).
  - Repository write permissions: the workflow needs to push tags / create releases (`contents: write` permission).
- **Workflow files** (suggested layout):
  - `.github/workflows/release.yml` &mdash; the orchestrating workflow. Inputs: version, deliverables to ship (`psgallery`, `nuget`, `github-release` checkboxes), `whatif`. Reuses the `microsoft/setup-msbuild@v2` + chocolatey-devpack setup from `build.yml`.
  - PSGallery step: invokes `Publish-VisioPSToGallery.ps1`. Already supports `-WhatIf`.
  - NuGet step: `nuget pack NuGet/VisioAutomation2010.nuspec` then `nuget push *.nupkg -Source https://api.nuget.org/v3/index.json -ApiKey $env:NUGET_API_KEY`.
  - GitHub Release step: `softprops/action-gh-release@v2` with the staged module folder zipped + the `.nupkg` as artifacts; auto-generated release notes from commits.
- **First-cut testing:**
  - Run with `-WhatIf` (PSGallery) / `--no-symbols --no-service-endpoint --skip-duplicate` (NuGet pack-only) / `dry_run: true` on the GitHub Release step to verify the workflow shape end-to-end without touching the public feeds.
  - First real run: probably a `4.6.2` patch with no behavior change (just to exercise the workflow), or wait for the next legitimate version bump.

#### Cross-refs

- Subsumes *Publish the PowerShell module to the PowerShell Gallery* and *Publish the NuGet package to nuget.org*. Builds on *Add CI*. NuGet release is gated on *Reconcile version numbers across artifacts* unless the workflow handles two version sources explicitly.

#### Effort

- M for PSGallery alone (the script does all the heavy lifting).
- +M for NuGet (no comparable script exists).
- +S for GitHub Release attachments (well-trodden action).
- Total: M&ndash;L depending on how many of the three are tackled in one go.

### Revise user-facing documentation for accuracy
- **What:** Audit the public gitbook docs ([VisioAutomation](https://saveenr.gitbook.io/visioautomation/) and [Visio PowerShell](https://saveenr.gitbook.io/visiopowershell/), source repo: [VisioAutomation_GitBook_Docs](https://github.com/saveenr/VisioAutomation_GitBook_Docs)) against the current API surface. Update or remove anything that no longer matches the code, and fill in coverage for cmdlets / APIs that have been added since the docs were last touched.
- **Why:** The docs have not been refreshed alongside recent changes; users hitting a stale example as their first impression is the worst kind of regression.
- **Approach (suggested):**
  - Start with the **PowerShell module** since it has the most cmdlet-by-cmdlet documentation surface and is the most user-facing.
  - For each cmdlet, verify it still exists, parameters still match, and the example still runs.
  - Do the C# library docs second.
  - Use the new [`docs/ARCHITECTURE.md`](ARCHITECTURE.md) and [`docs/GLOSSARY.md`](GLOSSARY.md) as the source of truth for terminology and structure.
- **Cross-refs:** Related to but distinct from *Decide where docs live long-term* — that item is about the gitbook-vs-in-repo *policy*; this item is about *accuracy of the existing user-facing content*.
- **Effort:** L (the cmdlet inventory alone is substantial).
