# Futures — 2026 Refresh Backlog

A running list of cleanup, modernization, and improvement items for the VisioAutomation solution. Items are grouped by theme. Each entry includes a one-line **What**, a **Why** (cost of leaving it), and a rough **Effort** (S / M / L). This is a *backlog* — items are not committed to or scheduled until pulled out into actual work.

---

## Roadmap (staged plan)

The 2026 refresh runs in three phases. Each backlog item below is tagged with its phase.

### Phase 1 — VS 2022 cleanup *(in progress)*
Stay on Visual Studio 2022 and the current TFMs (.NET Framework 4.5 for shipping libs). Code + docs improvements only, **no new features**. Anything that would destabilize a release (TFM jump, IDE jump, csproj-format change, breaking API change) waits for Phase 3.

Phase 1 items:
- *Revise user-facing documentation for accuracy* (the largest item)

Phase 1 items completed:
- ✅ *Fix the misnamed PowerShell loader script* — rewrote it to actually `Save-Module` from the PS Gallery
- ✅ *Add a `CLAUDE.md` at the repo root* — added with staged-plan, build commands, conventions, doc pointers
- ✅ *Update MSTest off the beta* — upgraded `MSTest.TestFramework` and `MSTest.TestAdapter` from `2.0.0-beta2` to `4.2.2`; bumped `VTest` TFM 4.5 → 4.7.2 to satisfy MSTest 4.x's floor
- ✅ *Add a per-project `README.md` for the larger projects* — `VisioAutomation/`, `VisioAutomation.Models/`, `VisioScripting/`, `VisioPowerShell/` (already had one)
- ✅ *Add a `CONTRIBUTING.md`* — covers branch, setup pointer, tests-need-Visio rule, code style, commits, changelog discipline, per-phase scope
- ✅ *Expand the root `readme.md`* — rewrote with pitch, install table, C# + PowerShell quick-start, doc links, license
- ✅ *Audit `Internal/` for dead code* — deleted orphaned `TempHelper.cs` + removed dead `InternalsVisibleTo("TestVisioAutomation")` attribute; spawned a follow-up item for misc warts found during the audit
- ✅ *Misc cleanups discovered during the Internal/ audit* (mostly) — moved misplaced `InternalsVisibleTo` attributes to `AssemblyInfo.cs`, deleted two orphaned VTest files, removed auto-generated `.sln.metaproj` from version control. `LinqExtensions` visibility-vs-folder mismatch deferred to Phase 3 as a breaking-namespace-change risk.
- ✅ *Add CI* (build-only) — `.github/workflows/build.yml` builds the solution on push/PR for `master` and `2026_Refresh`, pinned to VS 2022 MSBuild, NuGet packages cached. Test runs in CI deferred to Phase 3 (needs self-hosted runner with Visio).

### Phase 2 — Cut the final release
Tag and publish a final release of VisioAutomation (NuGet) and VisioPowerShell (PowerShell Gallery) with the refreshed docs. This is the demarcation line between the old-world (VS 2022 / .NET Framework 4.5 / current architecture) and the new-world. Existing consumers get one stable, well-documented release before the modernization changes land.

Phase 2 prerequisites (must be settled before the release ships):
- *Reconcile version numbers across artifacts* — needs a deeper conversation before a decision; **currently deferred**, do not implement until discussed.
- *Investigate flakiness from leftover Visio processes* — relevant to the release-verification flow we'll exercise in Phase 2.

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
- **Status:** Step 1 done. All shipping libraries (`VisioAutomation`, `VisioAutomation.Models`, `VisioScripting`, `VisioPowerShell`) and both sample projects (`VSamples`, `VSamples.Docs`) are now on **.NET Framework 4.5** — converged on the TFM `VisioPowerShell` was already using. Test projects intentionally left on their existing TFMs (`VTest` on 4.5; `VTest.Models` / `VTest.Scripting` / `VTest.PowerShell` on 4.7.2) since they don't ship.
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
- **Resolution:** [`.github/workflows/build.yml`](../.github/workflows/build.yml) added. Builds the solution in Debug on every push to `master` / `2026_Refresh` and on every PR. Pinned to VS 2022's MSBuild (`vs-version: '17.0'`) since VS 2026's MSBuild can't resolve the .NET Framework 4.5 reference assemblies the shipping libs need. NuGet packages are cached keyed on the hash of all `packages.config` files. Build status surfaces as a badge in the root README.
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
- **Implication:** the shipping fleet (currently on 4.5 after step 1 of *Consolidate target frameworks*) must bump again before VS 2026 can build it. Recommended landing point: **4.7.2** — clears the VS 2026 floor *and* converges with the existing test-project TFM in one move.
- **VisioPowerShell older-PowerShell support is preserved** by this bump: the older-PS floor is set by the `System.Management.Automation` v3 reference and the `ModuleToProcess`/`PowerShellVersion = 2.0` choices in [Visio.psd1](../VisioAutomation_2010/VisioPowerShell/Visio.psd1), not by the .NET Framework TFM. Bumping 4.5 → 4.7.2 doesn't change that.
- **Cross-refs:** Drives step 2 of *Consolidate target frameworks*. Supersedes *Decide whether to move to .NET 6/8* for now (defer that decision).
- **Effort:** S — bump TFMs, bump `VisualStudioVersion` in the .sln, open in VS 2026, full rebuild.

### Automate releases via GitHub CI — NuGet + PowerShell Gallery
- **What:** Replace the current manual release process with a GitHub Actions workflow that, on a tagged release, builds the solution, packs the NuGet package, packages the PowerShell module, and pushes to nuget.org and the PowerShell Gallery.
- **Why:** Manual releases are error-prone and infrequent. Automating them removes friction, makes versioning consistent, and means fixes can ship quickly. (User notes the PS module is believed to already exist at https://www.powershellgallery.com/packages/Visio — confirm ownership and credentials as a prerequisite.)
- **Subtasks:**
  - Confirm ownership of the `Visio` PowerShell Gallery package and the NuGet package identity.
  - Store API keys as GitHub repository secrets.
  - Define the release trigger (Git tag? Manual `workflow_dispatch`? GitHub Release?).
  - Decide on signing (Authenticode for the PS module DLLs?) before automating publish.
- **Cross-refs:** Subsumes *Publish the PowerShell module to the PowerShell Gallery* and *Publish the NuGet package to nuget.org*. Builds on *Add CI*. Depends on *Reconcile version numbers across artifacts* (need a single source of truth for the version a release stamps).
- **Effort:** M.

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
