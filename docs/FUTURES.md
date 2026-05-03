# Futures — 2026 Refresh Backlog

A running list of cleanup, modernization, and improvement items for the VisioAutomation solution. Items are grouped by theme. Each entry includes a one-line **What**, a **Why** (cost of leaving it), and a rough **Effort** (S / M / L). This is a *backlog* — items are not committed to or scheduled until pulled out into actual work.

---

## Roadmap (staged plan)

The 2026 refresh runs in three phases. Each backlog item below is tagged with its phase.

### Phase 1 — VS 2022 cleanup *(in progress)*
Stay on Visual Studio 2022 and the current TFMs (.NET Framework 4.5 for shipping libs). Code + docs improvements only, **no new features**. Anything that would destabilize a release (TFM jump, IDE jump, csproj-format change, breaking API change) waits for Phase 3.

Phase 1 items:
- *Update MSTest off the beta*
- *Audit `Internal/` for dead code*
- *Reconcile version numbers across artifacts*
- *Fix the misnamed PowerShell loader script*
- *Investigate flakiness from leftover Visio processes*
- *Add a `CLAUDE.md` at the repo root*
- *Add a `CONTRIBUTING.md`*
- *Expand the root `readme.md`*
- *Add a per-project `README.md` for the larger projects*
- *Revise user-facing documentation for accuracy* (the largest item)
- *Add CI* (build-only is enough for this phase)

### Phase 2 — Cut the final release
Tag and publish a final release of VisioAutomation (NuGet) and VisioPowerShell (PowerShell Gallery) with the refreshed docs. This is the demarcation line between the old-world (VS 2022 / .NET Framework 4.5 / current architecture) and the new-world. Existing consumers get one stable, well-documented release before the modernization changes land.

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
- **Step 2 (remaining):** Bump the shipping fleet again to clear the **VS 2026** floor (Framework 4.6.2 minimum). Recommended landing point: **4.7.2** — same TFM the test projects already use, so the whole solution converges on one number. See *Move development to Visual Studio 2026* below.
- **Why:** Mixed TFMs cause subtle binary-compatibility surprises (a test project on a higher TFM can use APIs the library under test cannot). Step 1 eliminated the production 4.0/4.5 split; step 2 will eliminate the 4.5/4.7.2 split between shipping libs and tests.
- **Effort:** S (already partially done).

### Migrate from `packages.config` to `PackageReference`
- **What:** Every csproj still uses the old `packages.config` NuGet model.
- **Why:** `PackageReference` is transitive, lockable, and the only model supported by `dotnet` CLI / SDK-style projects. Required before any modernization beyond Framework.
- **Effort:** S–M

### Update MSTest off the beta
- **What:** All test projects pin `MSTest.TestFramework` to `2.0.0-beta2`.
- **Why:** Pre-release dependency in test code is a smell; current MSTest is well past 3.x. Either bump to a current stable MSTest or migrate to xUnit/NUnit while we're touching it.
- **Effort:** S

### Add CI
- **What:** No GitHub Actions / Azure Pipelines configuration in the repo.
- **Why:** A simple build-only workflow would catch regressions immediately. The integration tests need a live Visio, so they would run on a self-hosted Windows runner with Visio installed (or be skipped/quarantined in CI).
- **Effort:** S for build-only; M for build + tests on a self-hosted runner.

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

### Audit `Internal/` for dead code
- **What:** The `VisioAutomation/Internal/` folder has accreted helpers over many years; some may be unused now.
- **Why:** Cleanup before any larger refactor.
- **Effort:** S–M.

---

## Tests

### Tests require a live Visio
- **What:** Every test project spins up a real Visio process via COM. There is no mock/fake layer.
- **Why (consider):** This is intentional — the library's whole job is to drive Visio, and mocking COM gives false confidence. But the lack of any non-Visio test surface means there's no quick `dotnet test` that runs anywhere. *Not necessarily a problem*, just worth a deliberate decision before adding CI.
- **Effort:** N/A — design decision, not a task.

### Investigate flakiness from leftover Visio processes
- **What:** Aborted test runs can leave Visio processes that lock files and break the next run.
- **Why:** Add a test-host shutdown hook or pre-run cleanup so re-runs are deterministic.
- **Effort:** S.

---

## Packaging & versioning

### Reconcile version numbers across artifacts
- **What:** The NuGet [`VisioAutomation2010.nuspec`](../NuGet/VisioAutomation2010.nuspec) is at `2.6.0`; the PowerShell [`Visio.psd1`](../VisioAutomation_2010/VisioPowerShell/Visio.psd1) is at `4.6.0`; csproj `AssemblyVersion`s are independent again.
- **Why:** Hard to tell at a glance which library version corresponds to which module version. Pick a single source of truth (e.g., a `Directory.Build.props` with one shared version) or document the versioning policy explicitly.
- **Effort:** S.

### Fix the misnamed PowerShell loader script
- **What:** [`DownloadFromPowerShellGallery.ps1`](../VisioAutomation_2010/VisioPowerShell/DownloadFromPowerShellGallery.ps1) does not download from the PowerShell Gallery — it `Import-Module`s the local `bin\Debug` build.
- **Why:** Misleading. Either rename it (e.g., `LoadFromBinDebug_Alt.ps1`) and delete if redundant with `LoadFromBinDebug.ps1`, or make it actually fetch from the Gallery.
- **Effort:** S.

### Publish the PowerShell module to the PowerShell Gallery
- **What:** The module is currently distributed only by manual install (`InstallForCurrentUser.ps1`).
- **Why:** Gallery publication makes `Install-Module Visio` work for users. Requires deciding on the publication identity, signing, and a release process.
- **Effort:** M — operational rather than coding work.

### Publish the NuGet package to nuget.org
- **What:** Same question for the NuGet package as for the PS module.
- **Effort:** S–M.

---

## Documentation

### Add a `CLAUDE.md` at the repo root
- **What:** Project-specific instructions for future Claude Code sessions: build commands, test rules (need Visio installed), where the public API lives, the `2026_Refresh` branch convention.
- **Why:** Loaded automatically into every Claude session in this repo; prevents re-discovering the same context next time.
- **Effort:** S.

### Add a `CONTRIBUTING.md`
- **What:** How to clone, build, run tests, the code style, the PR process.
- **Why:** Lowers the barrier for outside contributors and for the project's own future-self after another long pause.
- **Effort:** S.

### Expand the root `readme.md`
- **What:** Currently three lines. Add a short pitch, a code snippet, and links to `docs/OVERVIEW.md` and the gitbook user docs.
- **Why:** First impression for anyone landing on the GitHub page.
- **Effort:** S.

### Decide where docs live long-term
- **What:** User docs are in a separate repo on gitbook ([`VisioAutomation_GitBook_Docs`](https://github.com/saveenr/VisioAutomation_GitBook_Docs)); developer docs are now in `docs/` here.
- **Why:** Two-repo doc setups drift. Either keep them split with a clear policy (which doc lives where) or consolidate. No urgent action needed — just call out the policy in `OVERVIEW.md` once decided.
- **Effort:** S (policy) — or M (consolidation).

### Add a per-project `README.md` for the larger projects
- **What:** Short orientation file in `VisioAutomation/`, `VisioAutomation.Models/`, `VisioScripting/`, `VisioPowerShell/`.
- **Why:** When someone opens a single project in isolation (e.g., GitHub directory view), they see context immediately rather than having to navigate to `docs/`.
- **Effort:** S.

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
