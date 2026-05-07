# Futures — Releases & versioning

Backlog of items related to release process, version policy, and publishing to public package feeds. For the staged plan see [`../ROADMAP.md`](../ROADMAP.md). For what's already shipped see [`../COMPLETED.md`](../COMPLETED.md). Index of all backlog files: [`../FUTURES.md`](../FUTURES.md).

---

### Reconcile version numbers across artifacts *(decision: stay divergent for now; revisit at PS 5.1 manifest bump)*
- **What:** The NuGet [`VisioAutomation2010.nuspec`](../../NuGet/VisioAutomation2010.nuspec) is at `2.6.0`; the PowerShell [`Visio.psd1`](../../VisioAutomation_2010/VisioPowerShell/Visio.psd1) is at `4.7.2`; csproj `AssemblyVersion`s are independent again.
- **Why the question came up:** Hard to tell at a glance which library version corresponds to which module version. Same code (the PS module bundles the NuGet's DLLs) shipping under two different version numbers makes bug reports ambiguous and release coordination loose.
- **Decision (2026-05-08):** Stay divergent. NuGet continues on the `2.x` line, PS module on the `4.x` line. Re-evaluate convergence (or Option C: single source of truth via `Directory.Build.props` token substitution) when the [`Visio.psd1` deprecation warnings](#address-visiopsd1-deprecation-warnings-on-psgallery-publish) item is picked up &mdash; that's the manifest's `PowerShellVersion '2.0'` &rarr; `'5.1'` bump, which is itself a forcing function for re-touching the manifest. Reasoning: at this point the divergence is established, contributors and downstream consumers reason about the two version lines independently, and merging them mid-flight risks confusion during a release cycle. The PS-5.1 manifest bump is a natural integration point because it is already a "we are revisiting Visio.psd1's declared metadata" event, so adding a version-policy decision to the same change is low marginal cost.
- **History of options considered:**
  - **A &mdash; Converge:** both artifacts ship at the same number going forward. Rejected for now; revisit at the PS-5.1 bump.
  - **B &mdash; Document the divergence policy:** keep them independent, write down the rule. **Chosen for now** &mdash; this section *is* the documentation.
  - **C &mdash; Single technical source of truth:** `Directory.Build.props` + token substitution into nuspec/psd1. Better suited now that csprojs are SDK-style (Phase 3 SDK migration done 2026-05-06). Re-evaluate at the PS-5.1 bump.
- **Cross-refs:** [*Address `Visio.psd1` deprecation warnings*](#address-visiopsd1-deprecation-warnings-on-psgallery-publish) below &mdash; the convergence-revisit trigger. *Automate releases via GitHub CI* below already handles two version sources (each `release-*.yml` reads its own version from its own metadata file), so no workflow changes are required by this decision.
- **Effort to revisit:** S for the policy decision and doc updates; M if Option C is chosen at that point.

### Switch module-release builds from Debug to Release
- **What:** The release-prep script [`InstallForCurrentUser.ps1`](../../VisioAutomation_2010/VisioPowerShell/InstallForCurrentUser.ps1) hardcodes `$release = "Debug"` (line 69). The 4.6.1 release was published from the Debug build to keep the workflow unchanged, but for future releases we should ship the Release build — smaller binaries, no `DEBUG` symbols, no JIT debug overhead.
- **Why:** Shipping Debug builds to consumers is sloppy hygiene. Should be Release for any artifact that goes to a public feed (PSGallery, NuGet).
- **How:** Either flip the constant in `InstallForCurrentUser.ps1` (and document in the script comment that release-mode is now used for actual releases), or split the script into `InstallForCurrentUser.ps1` (Debug, dev convenience) and a separate `Stage-ReleaseBuild.ps1` (Release, used by `Publish-VisioPSToGallery.ps1`).
- **Cross-refs:** *Automate releases via GitHub CI* below — the CI workflow either flips the constant or stages the release config separately.
- **Effort:** S.

### Address `Visio.psd1` deprecation warnings on PSGallery publish
- **Status (2026-05-06):** the `CmdletsToExport` half landed (manifest now lists 64 cmdlets explicitly; the publish-time best-practice warning will be silent on the next release). The `ModuleToProcess` &rarr; `RootModule` rename and `PowerShellVersion` bump are still pending; deferred deliberately, not blocked. Customer-impact analysis below.
- **What's left:** The 4.7.0 publish to PSGallery (2026-05-06) emitted these still-active warnings from `Publish-Module`:
  - `The module manifest member 'ModuleToProcess' has been deprecated. Use the 'RootModule' member instead.` (fires twice &mdash; once during local staging, once on the gallery upload).
- **Why the warning fires:** [`Visio.psd1`](../../VisioAutomation_2010/VisioPowerShell/Visio.psd1) line 8 sets `ModuleToProcess = 'VisioPS.dll'` with an inline comment ("Use ModuleToProcess instead of RootModule because it works for both PowerShell 2.0 and 3.0") and the manifest declares `PowerShellVersion = '2.0'`. The compatibility target is essentially dead: PS 2.0 was removed from Windows 11; PSGallery's `Install-Module` requires PS 5.1+; [`Publish-VisioPSToGallery.ps1`](../../VisioAutomation_2010/VisioPowerShell/Publish-VisioPSToGallery.ps1) itself refuses to run below PS 5.1.
- **How to apply when picked up:**
  - Switch `ModuleToProcess` to `RootModule` in [`Visio.psd1`](../../VisioAutomation_2010/VisioPowerShell/Visio.psd1) (toggle the comment on lines 7-8). Bump `PowerShellVersion = '2.0'` to `'5.1'` to match the publish script's own minimum check. Optionally drop `CLRVersion = '4.0'` (paired with .NET Framework 4.0; the shipping libs target 4.5.2).

#### Customer impact of the `PowerShellVersion = '2.0'` &rarr; `'5.1'` bump

(Pre-derived 2026-05-06 so it doesn't have to be reasoned through again when this is picked up.)

| PS version | Ships with | Status | Impact of bump |
|---|---|---|---|
| 7.x | separate install, cross-platform | current | unaffected (7 satisfies the `>= 5.1` floor) |
| 5.1 | Windows 10 1607+ (2016), Win 11, Server 2016+ | current | **unaffected** &mdash; this is the de facto floor today |
| 5.0 | Windows 10 1507/1511 only (2015-2016) | unsupported Windows | newly blocked, but Windows itself is end-of-life |
| 4.0 | Windows 8.1, Server 2012 R2 | OOS since 2023 | newly blocked, but unlikely to be a VisioPS user |
| 3.0 | Windows 8 | OOS since 2023 | same |
| 2.0 | Windows 7 | OOS since 2020 | already broken &mdash; the binary cmdlets are compiled against `Microsoft.PowerShell.3.ReferenceAssemblies` and won't load on PS 2.0 regardless of what the manifest claims |

Effective customer impact: **zero**. The current `PowerShellVersion = '2.0'` is aspirational, not real; the binary already won't run on anything below PS 3.0, and the entire installation path (PSGallery + `Publish-VisioPSToGallery.ps1`) gates at 5.1. Bumping the manifest aligns the declaration with what the module actually requires; no extant user becomes unable to run it. LTSB 2016 (the ongoing compat constraint until 2026-10-13) ships with PS 5.1, so the bump does not affect those users either.

#### Maintenance note on `CmdletsToExport`
The list now has 64 entries. Adding a new cmdlet requires also adding it to the manifest. Two options to enforce when convenient:
- (a) Manual reminder in cmdlet-author docs / `CONTRIBUTING.md`.
- (b) Pre-publish check that loads the staged module, diffs `Get-Module -Name Visio | Select-Object -ExpandProperty ExportedCmdlets` against the manifest's `CmdletsToExport`, and fails if they differ. Drop into `Publish-VisioPSToGallery.ps1` between the staging step and the `Publish-Module` call.

Option (b) is a few lines of PowerShell; worth adding when the *PSGallery publish via "release first..."* item below lands, since that's the natural place to put the check.

- **Cross-refs:** *Automate releases via GitHub CI* below &mdash; the publish workflow should also surface these warnings as a CI signal so future drift is caught early.
- **Effort to finish:** S. Single commit, no behavior change for any extant user. The `RootModule` swap is one line; the `PowerShellVersion` bump is one line.

### Automate releases via GitHub CI *(in progress)*
- **What:** Replace the current manual release process with GitHub Actions workflows that handle **three deliverables** end-to-end:
  1. **PSGallery publish** of the `Visio` PowerShell module &mdash; ✅ landed, see below.
  2. **nuget.org publish** of the `VisioAutomation2010` NuGet package &mdash; ✅ landed, see below.
  3. **GitHub Release** with the built binaries (DLLs / `.zip` / the `.nupkg`) attached as downloadable artifacts &mdash; ✅ landed, see below.
- **Why:** Manual releases are error-prone and infrequent. The 4.6.1 release surfaced several PS-5.1 / PowerShellGet gotchas that are now baked into [`Publish-VisioPSToGallery.ps1`](../../VisioAutomation_2010/VisioPowerShell/Publish-VisioPSToGallery.ps1) and the [Publishing doc](https://saveenr.gitbook.io/visiopowershell/developer-info/publishing-to-powershell-gallery); automating those steps ensures future releases inherit the workarounds. GitHub Releases also give consumers a stable URL to download a specific version's binaries even if PSGallery / nuget.org are slow to update.

#### Status (2026-05-06)

**GitHub Release workflows** &mdash; both shipped 2026-05-04:
- [`.github/workflows/release-nuget.yml`](../../.github/workflows/release-nuget.yml) &mdash; reads `<version>` from [`NuGet/VisioAutomation2010.nuspec`](../../NuGet/VisioAutomation2010.nuspec), builds Debug, packs the `.nupkg`, builds a separate raw-DLL zip (same DLL list as the nuspec `<files>` group), creates a tag `VisioAutomation_<version>` and a GitHub Release with both attached.
- [`.github/workflows/release-psmodule.yml`](../../.github/workflows/release-psmodule.yml) &mdash; reads `ModuleVersion` from [`Visio.psd1`](../../VisioAutomation_2010/VisioPowerShell/Visio.psd1), builds Debug, stages `VisioPowerShell/bin/Debug` into a zip, creates a tag `VisioPS_<version>` and a GitHub Release with the zip attached.

Both have a `dry_run` input that builds artifacts and uploads them as workflow artifacts but skips tag/release creation. Both refuse to run if the derived tag already exists. Both share the `microsoft/setup-msbuild@v2` + NuGet-cache setup from [`build.yml`](../../.github/workflows/build.yml). Build is Debug to match how 4.6.1 shipped &mdash; switching to Release is still tracked in *Switch module-release builds from Debug to Release* above; tackle that separately.

**PSGallery publish workflow** &mdash; shipped 2026-05-06. [`.github/workflows/publish-psmodule.yml`](../../.github/workflows/publish-psmodule.yml) is `workflow_dispatch`-triggered, takes the release tag as input, downloads the `Visio-<version>.zip` artifact from the GH Release, extracts it, verifies the staged `ModuleVersion` matches the tag, runs a `CmdletsToExport` drift check (the `*` &rarr; explicit list change in `bd627062` is the prereq that makes this possible), forces TLS 1.2, runs `Publish-Module`, then verifies via `Find-Module` with retry. Has a `dry_run` input that exercises every step except the upload. Companion change: [`Publish-VisioPSToGallery.ps1`](../../VisioAutomation_2010/VisioPowerShell/Publish-VisioPSToGallery.ps1)'s tag step refactored to be idempotent so the local script and the workflow coexist for fallback / out-of-band publishes. The new flow shape is "release first, then publish from the release artifact" &mdash; the binary on PSGallery is exactly the one humans inspected on the GH Release. Setup required: add `PSGALLERY_API_KEY` as a repository secret.

**nuget.org publish workflow** &mdash; shipped 2026-05-08. [`.github/workflows/publish-nuget.yml`](../../.github/workflows/publish-nuget.yml) is `workflow_dispatch`-triggered, takes the release tag (e.g. `VisioAutomation_2.6.0`) as input, downloads the `VisioAutomation2010.<version>.nupkg` from the GH Release, opens it as a zip and verifies the embedded `.nuspec`'s `<id>` and `<version>` match the tag, then runs `dotnet nuget push` with `--skip-duplicate` (idempotent on re-runs) and verifies via nuget.org's V3 flat-container index with retry. Has a `dry_run` input that exercises every step except the upload. No equivalent of `publish-psmodule.yml`'s TLS-1.2 step or `CmdletsToExport` drift check is needed (PS 7 / `dotnet` handle TLS, and there's no per-file export manifest). Verification has a longer retry window (~5 min vs. PSGallery's ~30 s) because nuget.org's flat-container indexing lags 1-3 min behind the push and full validation can take 15+ min &mdash; the workflow surfaces a clear "listing may still be propagating; re-run is safe" message if the cap is hit. Setup required: add `NUGET_API_KEY` as a repository secret with push permission on the `VisioAutomation2010` package.

#### References for the workflow content

- **PSGallery publish** &mdash; [`Publish-VisioPSToGallery.ps1`](../../VisioAutomation_2010/VisioPowerShell/Publish-VisioPSToGallery.ps1) is the canonical battle-tested flow (TLS 1.2, `-Path` not `-Name`, `-ErrorAction Stop`, post-publish verification via `Find-Module`, then tag). It's callable from the workflow as-is; reads the API key from `$env:PSGalleryApiKey` or `-ApiKey`.
- **NuGet publish** &mdash; the package metadata is in [`NuGet/VisioAutomation2010.nuspec`](../../NuGet/VisioAutomation2010.nuspec) (currently `2.6.0`). No equivalent battle-tested script exists; the workflow needs a `nuget pack` + `nuget push` step or the equivalent `dotnet nuget push`. NuGet's gallery and `Publish-Module` don't share infrastructure &mdash; expect different gotchas.
- **GitHub Release** &mdash; the [`softprops/action-gh-release@v2`](https://github.com/softprops/action-gh-release) action handles upload-on-tag-push idiomatically. Alternative: `gh release create`.
- **Existing CI infrastructure** &mdash; [`.github/workflows/build.yml`](../../.github/workflows/build.yml) is the build-only workflow. Pinned to VS 2022 MSBuild via `microsoft/setup-msbuild@v2`. The release workflow should copy this setup verbatim.

#### Decisions to make first

- **One workflow or three?** Cleanest: a single `release.yml` triggered on `workflow_dispatch` (or tag) with conditional steps based on which deliverable to ship. Feature-flagged by inputs makes the first-cut testing easier.
- **Trigger.** Three reasonable choices:
  - `workflow_dispatch` with version + flags as inputs (manual, fully controlled, recommended for first cut).
  - Tag-push: `VisioPS_*` &rarr; PSGallery + GitHub Release for the module bundle; `VisioAutomation_*` &rarr; NuGet + GitHub Release for the library. Two separate workflow files keyed on tag pattern.
  - GitHub Release creation as the trigger (`on: release: types: [published]`). Less appealing because the manual `Publish-VisioPSToGallery.ps1` script already creates the tag at the end of a successful publish; reproducing the same ordering means the workflow creates the GitHub Release at the end too.
- **Tag-then-publish vs. publish-then-tag.** The 4.6.1 manual flow tagged **after** verifying the publish landed. Reproducing that ordering in CI pushes toward `workflow_dispatch` (publish, then tag from inside the workflow) rather than tag-push. Note: a subsequent GitHub Release creation step would then attach to that tag.
- **What artifacts go into the GitHub Release?** Candidates: the staged module folder zipped (the same content that's published to PSGallery), the `.nupkg` from the NuGet publish, possibly a separate "binaries-only" zip of the DLLs for users who don't want either package manager. Keep it small to start; one zip with the module is sufficient as a v1.
- **Build configuration.** Phase 1 shipped 4.6.1 from a Debug build (`InstallForCurrentUser.ps1` hardcodes `Debug`). Future releases should switch to Release; tracked in [the *Switch module-release builds from Debug to Release* item above](#switch-module-release-builds-from-debug-to-release). The CI workflow either flips the constant or stages the release config separately.
- **Signing.** Authenticode signing of the bundled DLLs is open. Required by neither PSGallery nor nuget.org but would silence the "publisher unknown" warning. Defer until the workflow is otherwise stable.
- **Version policy.** Module is at `4.6.1`; NuGet is at `2.6.0`. Until [*Reconcile version numbers across artifacts*](#reconcile-version-numbers-across-artifacts-phase-2-prereq--deferred-needs-discussion) is settled, the workflow has to handle two different version sources (read PS module version from `Visio.psd1`, NuGet version from `VisioAutomation2010.nuspec`). That's fine; just be explicit about it.

#### Subtasks

- **Confirm credentials and ownership:**
  - PSGallery: `Visio` package, key stored as GitHub secret (suggested name: `PSGALLERY_API_KEY`).
  - nuget.org: `VisioAutomation2010` package &mdash; confirm ownership and add the secret (suggested name: `NUGET_API_KEY`).
  - Repository write permissions: the workflow needs to push tags / create releases (`contents: write` permission).
- **Workflow files** (suggested layout):
  - `.github/workflows/release.yml` &mdash; the orchestrating workflow. Inputs: version, deliverables to ship (`psgallery`, `nuget`, `github-release` checkboxes), `whatif`. Reuses the `microsoft/setup-msbuild@v2` setup from `build.yml`.
  - PSGallery step: invokes `Publish-VisioPSToGallery.ps1`. Already supports `-WhatIf`.
  - NuGet step: `nuget pack NuGet/VisioAutomation2010.nuspec` then `nuget push *.nupkg -Source https://api.nuget.org/v3/index.json -ApiKey $env:NUGET_API_KEY`.
  - GitHub Release step: `softprops/action-gh-release@v2` with the staged module folder zipped + the `.nupkg` as artifacts; auto-generated release notes from commits.
- **First-cut testing:**
  - Run with `-WhatIf` (PSGallery) / `--no-symbols --no-service-endpoint --skip-duplicate` (NuGet pack-only) / `dry_run: true` on the GitHub Release step to verify the workflow shape end-to-end without touching the public feeds.
  - First real run: probably a `4.6.2` patch with no behavior change (just to exercise the workflow), or wait for the next legitimate version bump.

#### Cross-refs

- *Reconcile version numbers across artifacts* above &mdash; gates the NuGet release unless the workflow handles two version sources explicitly.
- *Switch module-release builds from Debug to Release* above &mdash; the CI workflow either flips the constant or stages the release config separately.

#### Effort

- M for PSGallery alone (the script does all the heavy lifting).
- +M for NuGet (no comparable script exists).
- +S for GitHub Release attachments (well-trodden action, already done).
- Total: M&ndash;L depending on how many of the three are tackled in one go.
