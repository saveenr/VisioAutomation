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

### Move to C# 14 / .NET 10
- **What:** Migrate the whole solution off .NET Framework to **.NET 10** (`net10.0-windows`, since the codebase depends on COM interop), and adopt **C# 14** as the language version. Replaces the older "Decide whether to move to .NET 6/8" item with a concrete landing point.
- **Why:** Long-term viability (.NET Framework gets security updates only) and access to a decade of language and runtime improvements. C# 14's extension-member support (see below) is the headline payoff for *this* codebase specifically &mdash; the helper layer is built around extensions on COM types, and the new feature reshapes a chunk of public API.

#### C# 14 features worth using
- **Extension members (extension types).** The big one. Today C# only supports extension *methods*, so every helper on a COM type has to be a method call: `shape.GetBoundingBox()`, `shape.GetXYFromPage(p)`, `page.DrawRectangle(rect)`. C# 14 introduces extension blocks that can carry extension *properties*, *indexers*, *static methods*, and instance methods on a per-type basis. Concretely, it lets the existing surface read more naturally:
  - `shape.GetBoundingBox()` &rarr; `shape.BoundingBox` (extension property).
  - The static `ShapeIDPairs.FromShapes(s1, s2)` helpers could become extension static methods on `IEnumerable<IVisio.Shape>`.
  - The drawing primitives in [`Extensions/PageExtensions.cs`](../../VisioAutomation_2010/VisioAutomation/Extensions/) could be reorganized into a single extension block on `IVisio.Page`.
  - Touches a substantial chunk of public surface: see [`Extensions/`](../../VisioAutomation_2010/VisioAutomation/Extensions/) (16 files, ~30 classes per the recently-rewritten [extension-methods gitbook page](https://saveenr.gitbook.io/visioautomation/extension-methods)). Any reshape here is a coordinated docs change as well; expect to rewrite the extension-methods page.
- **`field` keyword in auto-properties.** Cleans up the Cells records (`CustomPropertyCells`, `UserDefinedCellCells`, `ShapeFormatCells`, etc.) where today's pattern is `public Core.CellValue Value { get; set; }` and the `EncodeValues()` and similar methods reach across via property accessors. With `field`, validation / lazy-init / encoding logic can live inline without the verbosity of an explicit backing field.
- **Null-conditional assignment** (`x?.Y = z;`). Minor ergonomics; useful in a few `Cells` mutation paths where the target may be uninitialised.
- **`params Span<T>`.** Performance polish for variadics &mdash; e.g. `ShapeIDPairs.FromShapes(params IVisio.Shape[] shapes)` could become `params ReadOnlySpan<IVisio.Shape>` for callers that already have a span. Marginal in the COM-interop hot path; not load-bearing.
- **Lambda parameter modifiers** (`ref`/`out`/`in`). Niche; no obvious uses in this codebase.

#### Implications for `VisioPowerShell` (the binary PS module)
The thorniest piece. Today the module loads in **Windows PowerShell 5.1** (the Windows-default PS host, built on .NET Framework 4.x) and **PowerShell 7+** (built on modern .NET). Moving the module to .NET 10 forces a choice:

- **Option A &mdash; drop Windows PowerShell 5.1.** Module targets `net10.0-windows` only. PS 7+ users continue to work; PS 5.1 users get *"the specified module could not be loaded"*. Simpler build, simpler module, no maintenance overhead. But: PS 5.1 is bundled with every Windows install; PS 7 is a separate manual install. The audience overlap with VisioPS users (Windows desktop with Visio installed) skews toward the PS 5.1 default. Telemetry would help; absent that, expect a non-trivial fraction of users on 5.1.
- **Option B &mdash; multi-target the module.** csproj targets both `net472` (or `net48`) and `net10.0-windows`; the .psd1 manifest's `RootModule` becomes a tiny `.psm1` script that detects `$PSEdition` at load time and dot-sources the matching DLL. Two builds to publish, two sets of binaries in the module folder, but PS 5.1 keeps working. Maintenance overhead is real but bounded; mostly a one-time setup cost.
- **Option C &mdash; split into two modules.** `Visio` (PS 5.1 / .NET Framework, frozen at the time of the split) and `Visio7` or similar (PS 7+ / .NET 10, gets all new features). Confusing for users, doubles the docs surface. Probably not worth it.

Recommended path: **Option B** initially, plus a deprecation timeline for PS 5.1 support (tied to Microsoft's own PS 5.1 deprecation, currently informal but trending). Re-evaluate after 1-2 release cycles based on download split.

#### Other implications
- **`Microsoft.Office.Interop.Visio` PIA.** Works on modern .NET via the [`Microsoft.Office.Interop.Visio`](https://www.nuget.org/packages/Microsoft.Office.Interop.Visio/) NuGet (already used here). Embedded interop types (`<EmbedInteropTypes>true</EmbedInteropTypes>`) work on modern .NET but with caveats around runtime callable wrappers; expect to keep `false` (the current setting per [`Directory.Build.props`](../../VisioAutomation_2010/Directory.Build.props)).
- **NuGet package layout.** The `VisioAutomation2010.nuspec` `<files>` section currently puts DLLs in `lib/net452/`. Modern .NET means a new TFM in the package &mdash; either swap to `lib/net10.0-windows/` (drops .NET Framework support) or multi-target. The existing nuspec `<frameworkAssembly>` reference to the Visio PIA is already redundant (tracked separately under *Small follow-ups deferred from the migration*).
- **Globalization.** Modern .NET defaults to ICU for sorting / casing / date formatting; .NET Framework uses NLS. The `DATETIME(...)` formatting in `CustomPropertyCells` and the locale-formatted dates the characterization tests assert against (e.g. `3/31/2017 2:05:06 PM`) may behave differently. Re-run the test suite under .NET 10 and lock in any new formats; document the locale dependency that already exists.
- **Trimming / AOT.** COM interop is hostile to both. Don't enable; the runtime size benefit is tiny next to the COM surface anyway.
- **Test infrastructure.** VTest projects (currently `net472`) migrate alongside. MSTest 4.x already supports modern .NET; the `[AssemblyCleanup]` orphan-prevention in `Framework.AssemblyHooks` carries over.
- **Dev environment.** Needs the .NET 10 SDK installed locally and on the build runner. VS 2026 supports .NET 10 with the right workloads. Phase 3's SDK-style csproj migration already cleared the way; no further build-system surgery needed.
- **Version numbers.** This is a major version break for both artifacts. Natural breakpoints: `VisioAutomation2010` 2.x &rarr; 3.0; `Visio` PowerShell 4.x &rarr; 5.0. Resolves the version-divergence question as a side effect (both jump majors at the same time).

#### Sequencing
1. *Move development to Visual Studio 2026* lands first (already queued).
2. *Consolidate target frameworks* step 2 (4.5.2 &rarr; 4.7.2) lands after the LTSB 2016 sunset on 2026-10-13. Both shipping libs and tests converge on 4.7.2.
3. **This item lands after step 2.** Migrate `net472` &rarr; `net10.0-windows` for the libraries; decide on Option A / B / C for the PS module; bump major versions on the artifacts; rewrite the extension-methods gitbook page.

#### Cross-refs
- *Move development to Visual Studio 2026* above &mdash; precondition.
- *Consolidate target frameworks* above &mdash; precondition; a 4.7.2 baseline simplifies the .NET 10 migration vs. jumping straight from 4.5.2.
- *Reconcile version numbers across artifacts* in [`releases.md`](releases.md#reconcile-version-numbers-across-artifacts-phase-2-prereq--deferred-needs-discussion) &mdash; this item resolves it as a side effect (both artifacts hit the same major bump).
- *Run tests in CI* above &mdash; needed before this lands so the .NET 10 migration can be validated against the live test suite.

#### Effort
- L &mdash; touches every project, every test, the NuGet packaging, the PS module manifest. Plus a docs rewrite for the extension-methods page if extension members are adopted.
- The actual TFM swap is small per-csproj; the elapsed-time costs come from (a) the PS-module-compat decision, (b) audit-running the test suite under modern .NET to find globalization / interop drift, and (c) the public-API reshape if extension members are adopted.

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
