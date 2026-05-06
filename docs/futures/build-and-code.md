# Futures — Build, tooling, code & architecture

Backlog of build-system, tooling, and code/architecture items. For the staged plan and what's already shipped see [`../ROADMAP.md`](../ROADMAP.md) and [`../COMPLETED.md`](../COMPLETED.md). Index of all backlog files: [`../FUTURES.md`](../FUTURES.md).

---

## Build & tooling

### Consolidate target frameworks
- **Status:** Step 1 done. All shipping libraries (`VisioAutomation`, `VisioAutomation.Models`, `VisioScripting`, `VisioPowerShell`) and both sample projects (`VSamples`, `VSamples.Docs`) are now on **.NET Framework 4.5.2** (originally bumped from 4.0 → 4.5 in commit `2fd6b466`, then 4.5 → 4.5.2 to satisfy the available Developer Pack — see BUILDING.md). Test projects on **.NET Framework 4.7.2** (VTest moved there as part of the MSTest 4.x upgrade; the others were already there).
- **Step 2 (remaining):** Bump the shipping fleet again to clear the **VS 2026** floor (Framework 4.6.2 minimum). Recommended landing point: **4.7.2** — same TFM the test projects already use, so the whole solution converges on one number.
- **Why:** Mixed TFMs cause subtle binary-compatibility surprises (a test project on a higher TFM can use APIs the library under test cannot). Step 1 eliminated the production 4.0/4.5 split; step 2 will eliminate the 4.5/4.7.2 split between shipping libs and tests.
- **Deferred until 2026-10-13** when Windows 10 LTSB 2016 leaves Extended Support; bumping earlier would block enterprise users on locked LTSB images. See `enterprise_compat_ltsb2016.md` in the project memory.
- **Cross-refs:** *Move development to Visual Studio 2026* below — drives this. The Phase 3 SDK migration already eliminated the 4.5.2 Developer Pack install requirement via `Microsoft.NETFramework.ReferenceAssemblies` packages, so the dev-pack pain is no longer a forcing function for this bump.
- **Effort:** S (already partially done).

### Run tests in CI
- **What:** [`.github/workflows/build.yml`](../../.github/workflows/build.yml) currently builds only — the test suite isn't exercised by CI. The orphan-leak fix in Phase 1 (see *Investigate flakiness from leftover Visio processes* in [`../COMPLETED.md`](../COMPLETED.md#investigate-flakiness-from-leftover-visio-processes)) is a prerequisite for re-runs to be idempotent, so re-running tests in CI is now feasible from a process-hygiene standpoint.
- **Why:** Without test runs in CI, regressions only surface on a developer's local machine or after release. The whole point of [the test-cleanup work that landed in Phase 1](../COMPLETED.md#investigate-flakiness-from-leftover-visio-processes) was to make the suite trustworthy enough to gate releases on.
- **Constraint:** Tests need Microsoft Visio installed on the runner. GitHub-hosted Windows runners don't have Visio, so this needs a **self-hosted Windows runner** with Visio installed (or some other arrangement for ephemeral Visio installs).
- **Cross-refs:** Should land before *Automate releases via GitHub CI* in [`releases.md`](releases.md#automate-releases-via-github-ci-in-progress) (a working test gate is the natural pre-publish check). The *Tests require a live Visio* design-decision item in [`tests.md`](tests.md#tests-require-a-live-visio) frames the constraint.
- **Effort:** M (provisioning the self-hosted runner is the bulk; wiring up the workflow is small).

### Move development to Visual Studio 2026
- **What:** Bump the solution from VS 2022 (`VisualStudioVersion = 17.0` in the .sln) to VS 2026. Stay on .NET Framework — do not migrate to modern .NET (Core).
- **Constraint discovered during research:** VS 2026 supports .NET Framework targets **4.6.2, 4.7, 4.7.1, 4.7.2, 4.8, 4.8.1** only. Framework 4.0 / 4.5 / 4.5.x / 4.6 / 4.6.1 are **not** supported targets in VS 2026. Source: [Visual Studio 2026 Compatibility](https://learn.microsoft.com/en-us/visualstudio/releases/2026/compatibility).
- **Implication:** the shipping fleet (currently on 4.5.2 after step 1 of *Consolidate target frameworks*) must bump again before VS 2026 can build it. Recommended landing point: **4.7.2** — clears the VS 2026 floor *and* converges with the existing test-project TFM in one move.
- **VisioPowerShell older-PowerShell support is preserved** by this bump: the older-PS floor is set by the `System.Management.Automation` v3 reference and the `ModuleToProcess`/`PowerShellVersion = 2.0` choices in [Visio.psd1](../../VisioAutomation_2010/VisioPowerShell/Visio.psd1), not by the .NET Framework TFM. Bumping 4.5 → 4.7.2 doesn't change that.
- **Cross-refs:** Drives step 2 of *Consolidate target frameworks* above. Supersedes *Decide whether to move to .NET 6/8* below for now (defer that decision).
- **Effort:** S — bump TFMs, bump `VisualStudioVersion` in the .sln, open in VS 2026, full rebuild.

---

## Code & architecture

### Consider migrating off Visio 2010 PIA
- **What:** All projects reference `Microsoft.Office.Interop.Visio` v14 (Visio 2010 PIA). Visio is now on a much newer version (16.x, with Visio for Microsoft 365).
- **Why:** The 2010 PIA still works at runtime against newer Visio versions, so this isn't urgent. But APIs added since 2010 are inaccessible without rebinding to a newer interop assembly. Decide whether to stay on 2010 (max compatibility) or move forward (access to newer features).
- **Effort:** M — touches every project; needs a compatibility decision.

### Decide whether to move to .NET 6/8 (out of .NET Framework)
- **What:** Whole solution is .NET Framework. Modern .NET supports COM interop on Windows.
- **Why:** Long-term viability — .NET Framework only gets security updates. But COM interop on modern .NET has its own quirks, and the PowerShell module bridge (Windows PowerShell 5.1 vs PowerShell 7) becomes a bigger decision.
- **Cross-refs:** *Move development to Visual Studio 2026* above supersedes this for now (do that first; defer the .NET 6/8 question).
- **Effort:** L — major undertaking.

### Make `CustomPropertyCells` values not require manual `EncodeValues()`
- **What:** `CustomPropertyCells.Value` is stored as a Visio formula, so a literal string written as `cp.Value = "testVal"` produces the formula `testVal` (no quotes), which Visio tries to evaluate as a name reference and fails. To get a string literal stored, the caller has to either pre-quote (`"\"testVal\""`) or call `cp.EncodeValues()` before writing. Surfaced by [#117](https://github.com/saveenr/VisioAutomation/issues/117); tracked in [#144](https://github.com/saveenr/VisioAutomation/issues/144).
- **Why:** The trap is silent &mdash; the property gets created, just with the wrong value (typically `0`). The PowerShell `Set-VisioCustomProperty` cmdlet sidesteps it by calling `EncodeValues()` internally, but model-level callers (`CustomPropertyHelper.Set`, the DOM `ShapeList`, the directed-graph render path) get no help. Docs were patched to call out the encoding step explicitly as the immediate fix; the API ergonomics are still a foot-gun.
- **Options on the table** (see [#144](https://github.com/saveenr/VisioAutomation/issues/144) for full details and the recommended path): auto-encode in `CustomPropertyHelper.Set` (the choke-point); or add a clearer factory like `CustomPropertyCells.FromString(...)`; or add an `Encoded` flag on the cells; or leave the API alone and lean on docs.
- **Effort:** S-M.

### Move `LinqExtensions` out of `Internal/` (or rename the folder)
- **What:** `LinqExtensions` lives at `VisioAutomation/Internal/Extensions/LinqExtensions.cs` but is `public` and consumed across the assembly boundary by `VisioAutomation.Models` (`ShapeList` calls its `NotOfType<T>` method). The `public` visibility is therefore correct; the **folder name** is misleading.
- **Why deferred from Phase 1:** Either fix is technically a breaking namespace change for any external code that happens to use the type. Phase 1 was code+docs cleanup only; namespace shifts belong with the broader Phase 3 modernization where breaking changes are acceptable.
- **Options:**
  - Move `LinqExtensions.cs` out of `Internal/` (e.g. into `Extensions/` proper). Namespace becomes `VisioAutomation.Extensions`.
  - Rename `Internal/Extensions/` to a non-`Internal/` folder. Same namespace effect.
- **Cross-refs:** Surfaced during the Phase 1 *Misc cleanups discovered during the Internal/ audit* work (see [`../COMPLETED.md`](../COMPLETED.md#misc-cleanups-discovered-during-the-internal-audit-mostly) for the rest of that audit's outcomes).
- **Effort:** S — single file move + namespace fix-up.
