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

### Borrow ideas from VisioBot3000 for VisioPS ergonomics
- **What:** [`VisioBot3000`](https://github.com/MikeShepard/VisioBot3000) (Mike Shepard, [`PSGallery`](https://www.powershellgallery.com/packages/VisioBot3000)) is a separately-maintained PowerShell module for Visio automation with a meaningfully different design from VisioPS. Worth a deliberate look for techniques to borrow into VisioPS, especially around scripted-diagram ergonomics.
- **Why:** VisioPS is a binary cmdlet module focused on low-level COM-surface coverage (cells, custom properties, hyperlinks, locks, text formatting). VisioBot3000 is a 100% script-based DSL focused on *composing diagrams quickly*. The two emphases are complementary, and VisioPS users who reach for diagram-composition workflows currently don't have the ergonomic shorthand VisioBot3000 offers.

#### Distinctive ideas in VisioBot3000

(Source material: the [`README`](https://github.com/MikeShepard/VisioBot3000), the two intro blog posts [Part 1](https://powershellstation.com/2016/04/28/introducing-visiobot3000-part-1-clark-kent/) / [Part 2](https://powershellstation.com/2016/04/29/introducing-visiobot3000-part-2-superman/), and the [`PSGallery listing`](https://www.powershellgallery.com/packages/VisioBot3000) which enumerates 44 exported functions.)

1. **Stencil + shape nickname registry.** `Register-VisioStencil -Name Servers -Path C:\temp\SERVER_U.vssx`, then `Register-VisioShape -Name WebServer -From Servers -MasterName 'Web Server'`. After registration the stencil and master can be referenced by their friendly names anywhere in the script. VisioPS today asks the user to keep `$master` variables in scope, threaded into every `New-VisioShape` call.
2. **Dynamic function generation on registration.** This is the most distinctive idea. Registering a shape *creates a PowerShell function with that name*. So after `Register-VisioShape -Name WebServer ...` the script can call `WebServer -name PrimaryServer` directly, no `New-VisioShape -master ...` boilerplate. PS function-table mutation (`Set-Item function:WebServer ...`) is the mechanism.
3. **Block-style nested syntax for containers.** Containers take a script block whose contents render inside the container's bounds:

    ```powershell
    New-VisioContainer -shape (Get-VisioShape Domain) -name MyDomain -contents {
        New-VisioShape -master WebServer -name PrimaryServer -x 5 -y 5
        New-VisioShape -master DBServer  -name SQL01         -x 5 -y 7
    }
    ```

    Combined with the dynamic-function trick above this reduces to:

    ```powershell
    Domain MyDomain {
        WebServer PrimaryServer
        DBServer SQL01
    }
    ```

4. **Relative positioning cursor.** A module-level "next position" cursor with helpers (`Set-NextShapePosition`, `Set-RelativePositionDirection Vertical`). Default places each new shape just to the right of the last; users can flip direction inside a block. Means most diagrams need no explicit `-x`/`-y`. VisioPS today requires either an explicit `Position` or wraps a layout step around the shape drops.
5. **Connector by nickname.** `New-VisioConnector -from PrimaryServer -to SQL01 -name SQL -color Red -Arrow`. References shapes by their registered name rather than by `$shape` reference. Makes connector-heavy scripts readable.
6. **Verb-prefixed aliases for the common cmdlets.** `Diagram` for `New-VisioDocument`, `Stencil` for `Register-VisioStencil`, `Shape` / `Container`. Short noun-form reads more like a description than imperative commands.
7. **`Convert-VisioObjectToPSObject`** &mdash; marshals a live COM object into a flat `PSCustomObject` for pipeline-friendly filtering / inspection. Niche, but a recurring pain point in VisioPS too.

#### What VisioPS does better today (preserve)

- **Lower-level coverage.** ShapeSheet cells, custom properties (per-Type behavior matrix, typed setters), hyperlinks, locks, control handles, connection points, text formatting, page-level cells, layout cells. VisioBot3000 doesn't try to cover this surface.
- **Strong typing via binary cmdlets.** Parameter binding, IntelliSense in script-pane editors, deterministic error messages. VisioBot3000's dynamic-function generation forfeits some of this (the per-shape function has no static signature).
- **Test coverage.** 17 PS-side tests today, run against a live Visio in a test-host singleton. VisioBot3000 has no comparable test infrastructure.
- **Underlying .NET library.** VisioPS rides on `VisioAutomation` / `VisioAutomation.Models`, both publicly consumable for non-PS callers. VisioBot3000 is PowerShell-only.

#### Adoption path

Treat as a series of incremental adoptions, not an all-at-once port:

- **Phase 1 &mdash; nickname registry as opt-in helpers.** New cmdlets: `Register-VisioStencilNickname`, `Register-VisioShapeNickname`, `Get-VisioRegisteredShape`. These store name &rarr; (stencil-doc, master-name) mappings in module-level state and provide convenience lookup. No dynamic functions, no DSL &mdash; just a name registry. Low-risk, easy to test, doesn't perturb existing scripts.
- **Phase 2 &mdash; block-style script-block parameter on `New-VisioContainer`.** Add a `-Contents { ... }` script block that runs in a context where the container's coordinate frame is implicit. Composes well with both the nickname registry and the existing `New-VisioShape` cmdlet.
- **Phase 3 &mdash; relative-positioning cursor.** New cmdlets `Set-VisioNextShapePosition`, `Set-VisioRelativePositionDirection`, plus auto-fill behavior on `New-VisioShape` when `-Position` is omitted *and* a cursor is active. Opt-in: existing scripts that pass `-Position` get the same behavior they always did.
- **Phase 4 &mdash; the dynamic-function trick** (the headline VisioBot3000 idea). Hybrid binary+script module: the existing binary `VisioPS.dll` keeps the cmdlets; a thin companion `.psm1` provides the dynamic-function generation that turns nicknames into callable functions. Substantial complexity (binary + script module bridging, function lifetime per import session, tab-completion behavior on auto-generated functions); save for last.

Phases 1-3 are straightforward additions to the binary cmdlet surface. Phase 4 is the architectural shift &mdash; by then we'd know from real script use whether it's worth the complexity.

#### Open research before adopting

- **Maintenance status of VisioBot3000.** PSGallery shows v1.1 from Jan 2018 (no newer); GitHub has recent commits to master per the repo page. Confirm what's live for users, what's stale, and whether the project is actively maintained before borrowing implementation specifics rather than just ideas.
- **License compatibility.** Confirm VisioBot3000's license (likely MIT). VisioPS is MIT.
- **PS host compatibility.** VisioBot3000 is script-only PS, so it loads in PS 5.1 and PS 7. Any borrowed binary-side helper has to track the *Move to C# 14 / .NET 10* item's PS-edition compat decision (see above).
- **Tab-completion / StrictMode / discoverability** for dynamically-generated functions. PS users running with `Set-StrictMode -Version 2` may not get clean tab-completion on functions that don't exist until a `Register-*` call runs.

#### Cross-refs

- *Move to C# 14 / .NET 10* above &mdash; the PS-edition compat decision there constrains what's safe to ship in the binary half of any hybrid module shape.
- *Decide whether to document `VisioScripting` as a public API* in [`docs.md`](docs.md) &mdash; tangentially related (any DSL borrowing would presumably go through `VisioScripting.Client` rather than re-implementing automation primitives).

#### Effort

- Phase 1 (nickname registry): S.
- Phase 2 (block-style container): S.
- Phase 3 (positioning cursor): S-M.
- Phase 4 (dynamic-function DSL): M-L. Most of the cost is in the binary+script module bridge; the PS-side function-table manipulation itself is small.

Total: M for phases 1-3 if pursued together; +M-L if Phase 4 is added.

### Borrow ideas from PSVA for VisioPS bulk-operation cmdlets
- **What:** [`PSVA`](https://github.com/jrich523/PSVA) (jrich523) is a small script-only PS module for Visio automation. Much narrower than VisioBot3000 (10 commits, 6 stars, demo-driven), so the value is in specific high-level helpers, not an architectural pattern.
- **Why:** Two of PSVA's demo scripts ([`VisioDemo.ps1`](https://github.com/jrich523/PSVA/blob/master/VisioDemo.ps1), [`StackDemo.ps1`](https://github.com/jrich523/PSVA/blob/master/StackDemo.ps1)) showcase patterns VisioPS doesn't currently expose ergonomically: pipeline-friendly bulk shape operations, distribute-along-a-line layouts, side-and-alignment-based shape decoration, layer visibility toggles. Worth lifting the *cmdlet shapes* from PSVA even though the underlying implementation in PSVA is "pure COM" PowerShell (no DLLs) and VisioPS would implement them as binary cmdlets over the existing `VisioAutomation` library.

#### Distinctive helpers in PSVA

(Source material: the [`README`](https://github.com/jrich523/PSVA), [`VisioDemo.ps1`](https://github.com/jrich523/PSVA/blob/master/VisioDemo.ps1), [`StackDemo.ps1`](https://github.com/jrich523/PSVA/blob/master/StackDemo.ps1).)

1. **`Set-visShapeDistribution`** &mdash; distributes a list of shapes evenly along a horizontal or vertical axis with explicit spacing and start coordinates:

    ```powershell
    Set-visShapeDistribution -shape $shareShapes -type Horizontal -space .5 -startX 2 -startY 3
    ```

   Useful for laying out an unknown-cardinality result set (each item from a `Get-WmiObject`, etc.) in a row.

2. **`Add-visShapeConnection` as a pipeline target** &mdash; star-topology connections in one line:

    ```powershell
    $connections = $shareShapes | Add-visShapeConnection -FromShape $servershape
    ```

   Each shape on the pipeline becomes a connector from the named source shape. Bulk wire-up.

3. **`Attach-visShape`** &mdash; attach a shape to one or more "base" shapes on a chosen side with an alignment policy:

    ```powershell
    Attach-visShape -Shape $top    -Selection ($s1,$s2) -Side Top    -Alignment Stretch
    Attach-visShape -Shape $bottom -Selection ($s1,$s2) -Side Bottom -Alignment RightOrBottom
    ```

   Sides: `Top` / `Bottom` / `Left` / `Right`. Alignments: `Stretch` / `LeftOrTop` / `RightOrBottom`. The "border-decorate this group of shapes" pattern is common for labelled boxes, sidebars, headers, footers; VisioPS has no direct cmdlet for it today. Note that PSVA's README acknowledges the "VERB warning for the use of Attach" &mdash; a `Set-`/`Add-` verb would replace it cleanly in a VisioPS port.

4. **Layer cmdlets** &mdash; `Add-visShapeToLayer` (pipe shapes onto a named layer, `-Force` creates the layer) and `Switch-visLayerVisibility` (toggle visibility for one or many layers by name):

    ```powershell
    $connLayer = $connections | Add-visShapeToLayer -Layer ConnLayer -Force
    Switch-visLayerVisibility $connLayer,"connector"
    ```

5. **`Set-visShapeAutoAlignment`** &mdash; auto-align selected shapes via Visio's built-in alignment feature.

#### Gap audit (completed 2026-05-07, [#150](https://github.com/saveenr/VisioAutomation/issues/150))

Each PSVA helper mapped against the current VisioPS cmdlet surface (64 cmdlets in `Visio.psd1` `CmdletsToExport` as of 4.7.2) and the underlying `VisioScripting.Client.*` command groups:

| PSVA helper | Coverage today | Classification |
|---|---|---|
| **`Set-visShapeDistribution`** (place N shapes on a horizontal/vertical line at explicit `-startX` / `-startY` with `-space N` spacing) | `Format-VisioShape -DistributeHorizontal` / `-DistributeVertical` exists, but it calls Visio's built-in `visCmdDistributeHSpace` / `visCmdDistributeVSpace` &mdash; *redistribute existing extents*, not place-at-coordinates. Underlying `VisioScripting.Arrange.DistributeOnAxis` has the same shape. | **Real gap.** Different operation than what's there today. |
| **`Add-visShapeConnection`** as pipeline target (`$shapes \| Add-visShapeConnection -FromShape $src` for star-topology bulk connect) | `Connect-VisioShape` exists ([VisioAutomation_2010/VisioPowerShell/Commands/VisioShape/ConnectVisioShape.cs:9-13](../../VisioAutomation_2010/VisioPowerShell/Commands/VisioShape/ConnectVisioShape.cs)). Both `From` and `To` are `IVisio.Shape[]` with `[Parameter(Position, Mandatory)]` &mdash; **no `ValueFromPipeline=true`**. | **Partial.** Function present, pipeline shape missing. Phase B fix: add a pipeline-friendly parameter set. |
| **`Attach-visShape`** (border-decorate: attach shape to one side of a group with Stretch / LeftOrTop / RightOrBottom alignment) | No equivalent. Neither cmdlet surface nor `VisioScripting.Client.*` has a side-attach primitive. Could in principle compose from `XForm*` cell writes, but there is no helper. | **Real gap.** Net-new cmdlet (verb correction: `Set-VisioShapeAttachment` or `Add-VisioShapeAttachment`). |
| **`Add-visShapeToLayer`** (`-Force` creates the named layer if missing; pipe shapes onto it) | No `*-VisioLayer` cmdlet at all. `VisioScripting.Client.Layer` exposes only `FindLayersOnPageByName` and `GetLayersOnPage` ([VisioAutomation_2010/VisioScripting/Commands/LayerCommands.cs:15-49](../../VisioAutomation_2010/VisioScripting/Commands/LayerCommands.cs)). No "create layer", no "add shape to layer", no visibility toggle. | **Real gap at both layers** (cmdlet surface *and* scripting library). VisioScripting.LayerCommands needs `CreateLayer`, `AddShapeToLayer`, `SetLayerVisibility` first; cmdlets are thin wrappers on top. |
| **`Switch-visLayerVisibility`** (toggle visibility for one or many layers by name) | Same as above &mdash; nothing on either layer of the stack. | **Real gap** (couples to the `Add-visShapeToLayer` work; same VisioScripting.LayerCommands extension would carry both). |
| **`Set-visShapeAutoAlignment`** (auto-align selected shapes via Visio's built-in alignment) | `Format-VisioShape -AlignHorizontal` (Left/Center/Right) and `-AlignVertical` (Top/Center/Bottom) exist ([VisioAutomation_2010/VisioPowerShell/Commands/VisioShape/FormatVisioShape.cs:24-29](../../VisioAutomation_2010/VisioPowerShell/Commands/VisioShape/FormatVisioShape.cs)) and call `VisioScripting.Arrange.AlignHorizontal`/`AlignVertical`, which use `Selection.Align(...)` directly. | **Already exists** &mdash; different cmdlet shape (omnibus `Format-VisioShape` rather than dedicated `Set-VisioShapeAutoAlignment`), same function. |

#### Audit summary

- **4 real gaps** worth lifting: positional distribution; side-attach; layer create/add; layer-visibility toggle. The two layer gaps couple together and need VisioScripting.LayerCommands extended first.
- **1 partial:** `Connect-VisioShape` needs a pipeline-friendly parameter set. Back-compat-safe Phase B work.
- **1 already covered:** `Set-visShapeAutoAlignment` &harr; `Format-VisioShape -Align*`. The May scoping review can decide whether the standalone cmdlet shape is worth duplicating for ergonomics or whether the omnibus is fine.

The "Adoption path" sub-section below is unchanged by these findings &mdash; Phase A is the four real gaps; Phase B is the `Connect-VisioShape` pipeline polish.

#### Adoption path

If the audit confirms gaps, treat as additive cmdlets, no architecture change:

- **Phase A &mdash; missing primitives:** add cmdlets for the four real gaps the audit confirmed: `Set-VisioShapeDistribution` (positional distribute with explicit start + spacing), `Add-VisioShapeAttachment` (the side-attach), `Add-VisioShapeToLayer` (with `-Force` to create the layer), `Set-VisioLayerVisibility` (toggle named layers). The two layer items need `VisioScripting.LayerCommands` extended with `CreateLayer` / `AddShapeToLayer` / `SetLayerVisibility` first; cmdlets are thin wrappers on top.
- **Phase B &mdash; pipeline-shape polish:** if any existing VisioPS cmdlet has the right *function* but the wrong *parameter shape* for pipeline use (e.g. `Connect-VisioShape` doesn't accept shapes from the pipeline), add `[Parameter(ValueFromPipeline=$true)]` overloads. Back-compat-safe.

#### Open research before adopting

- **Project activity.** PSVA has 10 total commits on master, no recent activity, no PSGallery listing observed. Treat as a pattern-mine, not a live dependency. Nothing to wait on or coordinate with the upstream.
- **License.** PSVA's repo doesn't show a `LICENSE` file in the top-level listing. Confirm before lifting any *implementation*; lifting cmdlet *shapes* and *parameter names* is fine on its own.
- **Existing cmdlet inventory.** ~~The audit step above is the prerequisite to filing concrete subtasks.~~ Audit completed 2026-05-07 ([#150](https://github.com/saveenr/VisioAutomation/issues/150)); see *Gap audit* table above.

#### Cross-refs

- *Borrow ideas from VisioBot3000 for VisioPS ergonomics* above &mdash; complementary direction. VisioBot3000 emphasises *composition* (DSL, nicknames, dynamic functions); PSVA emphasises *bulk operations on existing shapes* (distribute, attach, pipe-connect, layer-toggle). Both could be pursued in parallel; they don't conflict.

#### Effort

- ~~Half-day audit pass to map PSVA helpers against existing VisioPS / VisioScripting surface.~~ **Done 2026-05-07** ([#150](https://github.com/saveenr/VisioAutomation/issues/150)).
- S per cmdlet for the four additive gap-fillers (Phase A). The two layer cmdlets share an underlying `VisioScripting.LayerCommands` extension that's done once.
- S total for pipeline-shape polish (Phase B): one extra parameter set on `Connect-VisioShape`.

### Evaluate NetOffice / NetOfficeFw as a replacement for the Visio PIA
- **What:** [`NetOffice`](https://netoffice.io/) (originally Sebastian Lange) and its actively-maintained fork [`NetOfficeFw`](https://github.com/NetOfficeFw/NetOffice) (Jozef Izso) provide managed wrapper assemblies that replace the Microsoft Office Primary Interop Assemblies. The Visio binding ships as [`NetOfficeFw.Visio`](https://www.nuget.org/packages/NetOfficeFw.Visio) on NuGet. Two angles to consider: **use directly** as a replacement for our `Microsoft.Office.Interop.Visio` reference, or **learn from** even if we don't adopt &mdash; particularly the COM-proxy cleanup pattern.
- **Why:** The library is "syntactically and semantically identical to the Microsoft Interop Assemblies" (per [netoffice.io](https://netoffice.io/)) so migration is mostly mechanical, and it provides three practical wins over the bare PIA path:
  1. **No PIA deployment hurdle.** NetOffice is a regular managed assembly; no PIA registration, no Office-version-specific binding redirect dance.
  2. **Cross-version Visio support.** One assembly works against multiple Office / Visio versions. We currently bind to the Visio 2010 PIA (`Microsoft.Office.Interop.Visio` v14) and rely on it being forward-compatible at runtime against newer Visio installs (the *Consider migrating off Visio 2010 PIA* item below frames the choice). NetOffice handles multi-version support explicitly.
  3. **Automatic COM-proxy cleanup.** NetOffice's wrapper objects manage the `Marshal.ReleaseComObject` lifecycle for you. The current codebase has bespoke cleanup patterns (`Framework.VTestAppRef.QuitVisioApplication` swallows COMException during teardown; `ApplicationCommands.cs` does `documents.Close(true); app.Quit(true)` carefully). Adopting NetOffice would eliminate a class of orphan-process bugs we already chased in the test infrastructure (the 2026-05 orphan-leak fix in [`COMPLETED.md`](../COMPLETED.md)).

#### Maintenance status
- **NetOfficeFw (parent project):** active. v1.9.8 released February 2026; 754 commits on main; 44 total releases. MIT licensed. Recent release cadence is healthy.
- **`NetOfficeFw.Visio` NuGet:** last published version 1.8.1 (February 2021). Roughly 5 years between the Visio sub-package's last NuGet upload and the parent project's most recent release, which suggests the Visio binding is stable rather than abandoned, but worth confirming by checking the parent repo's `NetOffice.VisioApi` source for activity since 2021.
- **Original `NetOffice.*` packages:** legacy. NetOfficeFw is the canonical successor; the original packages should not be used for new work.

#### What's needed to evaluate "use directly"
A spike-grade evaluation would answer:
1. **Surface coverage.** Does `NetOfficeFw.Visio` expose the full set of Visio COM types VisioAutomation uses? Quick audit against [`VisioAutomation/Shapes/`](../../VisioAutomation_2010/VisioAutomation/Shapes/), [`VisioAutomation/Pages/`](../../VisioAutomation_2010/VisioAutomation/Pages/), and the rest of the helpers, looking for any IVisio.* type that might not have a NetOffice wrapper.
2. **Modern .NET targeting.** `NetOfficeFw.Visio` 1.8.1 targets `net40`-`net481` per the NuGet metadata (.NET Framework only). For the [*Move to C# 14 / .NET 10*](#move-to-c-14--net-10) item, NetOffice's modern-.NET support is a load-bearing input. Confirm whether the parent project ships modern-.NET TFMs (the README notes ".NET Framework 4.6 or higher"; would need to check the build to see what TFMs the actual binaries support).
3. **API translation cost.** NetOffice claims syntactic compatibility with the Microsoft PIA; verify by translating a representative slice of the codebase (e.g. [`CustomPropertyHelper.cs`](../../VisioAutomation_2010/VisioAutomation/Shapes/CustomPropertyHelper.cs)) and seeing what changes. Expected pattern: `using IVisio = Microsoft.Office.Interop.Visio;` becomes `using IVisio = NetOffice.VisioApi;`. If the change is one-line per file, the bulk migration is low-cost.
4. **Performance.** NetOffice wraps every COM call. VisioAutomation has high-throughput call paths (Section reads, batched ShapeSheet writes via `SrcWriter`/`SidSrcWriter`). Run the existing test suite under both stacks and compare wall-clock; the COM overhead is typically per-call, so the impact varies with call shape.
5. **Test-suite re-validation.** All four test projects re-run end-to-end. Particular attention to the singleton-Visio + `[AssemblyCleanup]` orphan-prevention pattern in [`Framework.AssemblyHooks`](../../VisioAutomation_2010/VTest/Framework/AssemblyHooks.cs) and equivalents &mdash; NetOffice's automatic-cleanup behavior may interact unexpectedly with our explicit `app.Quit(true)` calls.

#### What we could learn even if we don't adopt
- **COM-proxy cleanup pattern.** NetOffice's release-on-dispose model is more disciplined than what we do today. Could lift the pattern (an `IDisposable` wrapper that releases a referenced COM object on disposal) into a tiny internal helper, used at strategic points (test teardown, scripting `Client` lifetime).
- **Cross-version dispatch pattern.** How NetOffice handles "this method exists on Visio 16 but not 14" without breaking compile-time binding. We've never needed this because we bind to v14 and rely on forward-compat, but it's a useful pattern to have catalogued for any future Visio-version-aware features.
- **Diagnostic surface.** NetOffice surfaces "method invoke failed" errors with richer context (target, member, args) than the bare COMException we get from the PIA today. If we don't migrate, we could still copy the diagnostic-wrapping shape (similar to what [#144](https://github.com/saveenr/VisioAutomation/issues/144)'s detect-and-rethrow already does for the formula-error case).

#### Trade-offs against staying on the Visio PIA
- **Adopt NetOffice (use directly):** lose the official-Microsoft pedigree of the PIA, gain version-agility and cleanup hygiene. Risk: we depend on a third-party project; if NetOfficeFw.Visio falls behind on a future Visio API surface we need, we have no Microsoft-supported path to that surface without re-wiring back to the PIA.
- **Stay on the PIA (status quo):** lose nothing; gain nothing new. The current setup works.
- **Hybrid:** keep PIA for shipping libs; use NetOffice in tests only where the cleanup hygiene is most useful. Probably not worth the complexity unless the test suite specifically benefits.

#### Cross-refs
- *Move to C# 14 / .NET 10* above &mdash; NetOffice's modern-.NET support status (or lack of it) is a key input. If NetOfficeFw doesn't ship modern-.NET TFMs, we can't adopt it on the path to .NET 10 and would be stuck choosing between the modern Microsoft PIA package and continuing on .NET Framework.
- *Consider migrating off Visio 2010 PIA* above &mdash; NetOffice is one concrete answer to that item ("yes, leave the 2010 PIA, migrate to NetOffice's cross-version wrappers"). The two items resolve together.

#### Effort
- **Spike** (answer the five questions above): S-M. ~1 day of focused work; produces a go / no-go memo.
- **Full migration** if the spike says "go": M-L. Mechanical search-and-replace across ~30 csproj files for the namespace change, plus full test-suite re-validation. The bulk of the effort is in re-running tests and fixing whatever doesn't translate cleanly.
- **Pattern-mine without migrating:** S. Lift the COM-cleanup `IDisposable` pattern as an internal helper, ~half-day.

### Move `LinqExtensions` out of `Internal/` (or rename the folder)
- **What:** `LinqExtensions` lives at `VisioAutomation/Internal/Extensions/LinqExtensions.cs` but is `public` and consumed across the assembly boundary by `VisioAutomation.Models` (`ShapeList` calls its `NotOfType<T>` method). The `public` visibility is therefore correct; the **folder name** is misleading.
- **Why deferred from Phase 1:** Either fix is technically a breaking namespace change for any external code that happens to use the type. Phase 1 was code+docs cleanup only; namespace shifts belong with the broader Phase 3 modernization where breaking changes are acceptable.
- **Options:**
  - Move `LinqExtensions.cs` out of `Internal/` (e.g. into `Extensions/` proper). Namespace becomes `VisioAutomation.Extensions`.
  - Rename `Internal/Extensions/` to a non-`Internal/` folder. Same namespace effect.
- **Cross-refs:** Surfaced during the Phase 1 *Misc cleanups discovered during the Internal/ audit* work (see [`../COMPLETED.md`](../COMPLETED.md#misc-cleanups-discovered-during-the-internal-audit-mostly) for the rest of that audit's outcomes).
- **Effort:** S — single file move + namespace fix-up.
