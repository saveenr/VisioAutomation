# Completed — 2026 Refresh Done Work

A historical record of FUTURES.md items that have landed. The point is institutional memory: *what was done, why, what was tried, what didn't work*. This file is append-only — entries shouldn't be edited after landing except for typo or link fixes.

Forward-looking work lives in [`FUTURES.md`](FUTURES.md). The high-level "what shipped in Phase N" headline summary also lives in [`FUTURES.md`](FUTURES.md) (under each phase header); this file is the per-item detail behind those bullets.

When something on the FUTURES.md backlog completes, the Resolution paragraph moves here verbatim, the FUTURES.md body entry is deleted, and a one-line bullet is added to the relevant phase summary in FUTURES.md. See [`CONTRIBUTING.md`](../CONTRIBUTING.md) for the convention.

Items are grouped by **phase**, then by the same **category** they had in FUTURES.md (Build & tooling, Code & architecture, Tests, Documentation). Within a category they appear in the order they completed.

---

## Phase 1 — VS 2022 cleanup (done; merged to master 2026-05-03)

Headline summary in [`FUTURES.md`](FUTURES.md#phase-1--vs-2022-cleanup-done-merged-to-master-2026-05-03). Full detail per item below.

### Build & tooling

#### Update MSTest off the beta
- **What:** All test projects pinned `MSTest.TestFramework` and `MSTest.TestAdapter` to `2.0.0-beta2`.
- **Resolution:** Upgraded both packages to **4.2.2** (latest stable) across all four test projects. Required bumping VTest from .NET Framework 4.5 → 4.7.2 (MSTest 4.x's floor is 4.6.2); the other test projects were already on 4.7.2. Solution builds cleanly. Test code did not need any changes — MSTest's `[TestMethod]`/`Assert.*` API surface is stable across the version jump. Tests not actually run (need a live Visio).

#### Add CI (build-only)
- **Resolution:** [`.github/workflows/build.yml`](../.github/workflows/build.yml) added. Builds the solution in Debug on every push to `master` / `2026_Refresh` and on every PR. Pinned to VS 2022's MSBuild (`vs-version: '17.0'`) since VS 2026's MSBuild can't resolve the .NET Framework 4.5.2 reference assemblies the shipping libs need. The workflow also installs the .NET Framework 4.5.2 Developer Pack via chocolatey before building (those reference assemblies aren't on the runner image). NuGet packages are cached keyed on the hash of all `packages.config` files. Build status surfaces as a badge in the root README.
- **Tail (still active in FUTURES.md):** Running the tests themselves in CI requires a self-hosted Windows runner with Visio installed; tracked under *Run tests in CI* in [`FUTURES.md`](FUTURES.md).

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
- **Tail (still active in FUTURES.md):** `LinqExtensions` visibility-vs-folder mismatch is tracked under *Move `LinqExtensions` out of `Internal/` (or rename the folder)* in [`FUTURES.md`](FUTURES.md).

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
