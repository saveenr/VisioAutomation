# Futures — Dev team identity

Backlog of items related to the dev-team identity used across this codebase, its artifacts, and external services. For the staged plan see [`../ROADMAP.md`](../ROADMAP.md). For what's already shipped see [`../COMPLETED.md`](../COMPLETED.md). Index of all backlog files: [`../FUTURES.md`](../FUTURES.md).

---

### Transition dev team identity from "Saveen" to "SevenPens" *(in progress; multi-axis)*

- **Decision (2026-05-07):** development responsibilities are handing off from the historical `Saveen` / `saveenr` identity to the `SevenPens` identity going forward. The transition is multi-axis &mdash; each axis has its own cost, blocker, and reversibility profile, so they should be tracked and tackled independently rather than as one mass find-and-replace. Some axes are already done; some are easy when picked up; some require external coordination; one is a load-bearing legal question that should not be touched casually.

#### Axis 1: Commit author identity *(done)*
- All recent commits are authored by `TheSevenPens` per git config. Verified via `git log --format='%an'` on the recent stretch of master.
- No further action.

#### Axis 2: nuget.org publishing identity *(done 2026-05-07)*
- `SevenPens` is co-owner of [`VisioAutomation2010`](https://www.nuget.org/packages/VisioAutomation2010/) (saveenr remains co-owner for historical continuity). The GitHub repo's `NUGET_API_KEY` secret is generated under SevenPens. nuget.org's Microsoft-package compliance gate fires on uploader account, so SevenPens uploads pass cleanly while saveenr uploads fail. Detail in [`releases.md`](releases.md#microsoft-package-compliance-gate-on-the-saveenr-nugetorg-account-operational-quirk-discovered-2026-05-07-during-the-300-publish).
- No further action on this axis.

#### Axis 3: PSGallery publishing identity *(done 2026-05-07)*
- `SevenPens` is now co-owner of the [`Visio` PSGallery module](https://www.powershellgallery.com/packages/Visio); saveenr remains co-owner. `PSGALLERY_API_KEY` rotated to a SevenPens-generated key on 2026-05-07.
- **End-to-end validation deferred** to the next PSGallery release. Unlike axis 2, where the rotation was forced by an actual rejection on the saveenr key, axis 3 was preemptive &mdash; PSGallery hasn't tightened enforcement yet. So we know the rotation happened (secret timestamp confirms), but the new key hasn't been exercised against an upload. The next PSGallery release (whenever it ships) is the implicit smoke test.
- Memory rule covering both feeds: [`nuget_publish_identity.md`](../../../../.claude/projects/C--Users-savee-Documents-GitHub-VisioAutomation/memory/nuget_publish_identity.md) now consolidates the rule for both `NUGET_API_KEY` and `PSGALLERY_API_KEY` &mdash; same reasoning, same workaround.

#### Axis 4: Display authorship in artifact metadata *(done 2026-05-07)*
- All five displayed-author / copyright fields rewritten from `saveenr` / `Saveen Reddy` to `SevenPens`:
  - [`Visio.psd1`](../../VisioAutomation_2010/VisioPowerShell/Visio.psd1) `Author` and `Copyright`.
  - [`VisioAutomation2010.nuspec`](../../NuGet/VisioAutomation2010.nuspec) `<authors>` and `<copyright>`.
  - [`VPlayground/Properties/AssemblyInfo.cs`](../../VisioAutomation_2010/VPlayground/Properties/AssemblyInfo.cs) `AssemblyCopyright`.
- Both CHANGELOGs got an `[Unreleased]` "Changed" entry noting the brand swap. Picked up by the next NuGet release (already-published 3.0.0 retains the old `Copyright Saveen Reddy` line on its nuget.org page; that's frozen) and the next PSGallery release. Coupled to Axis 7 (handled together).
- **Adjacent inconsistency noticed but not addressed:** 10 of the 11 csproj `AssemblyInfo.cs` files have `AssemblyCopyright("Copyright © 2021")` (no name, just year). VPlayground was the only one with a person name. The "Copyright © 2021" form is an existing inconsistency unrelated to this identity transition; if the codebase wants to harmonize on a single canonical form (e.g., `"Copyright © 2026 SevenPens"`), that's its own one-shot pass &mdash; flag as a follow-up if desired.

#### Axis 5: Hosting URLs (`github.com/saveenr/...`, `saveenr.gitbook.io/...`) *(scheduled December 2026 via [milestone](https://github.com/saveenr/VisioAutomation/milestone/1); tracked in [#146](https://github.com/saveenr/VisioAutomation/issues/146) and [#147](https://github.com/saveenr/VisioAutomation/issues/147))*
- **Decision (2026-05-07):** do the migration with redirects from the old locations preserved. New canonical URLs live under SevenPens-owned hosting; old saveenr-prefixed URLs continue to resolve via redirects so external consumers don't break.
- **Three-phase plan, with phase 5a split into two GitHub issues for separate tracking:**
  - **Phase 5a-1 &mdash; transfer the GitHub repo.** Tracked in [#146](https://github.com/saveenr/VisioAutomation/issues/146) (milestone `2026-12`). GitHub auto-creates a permanent redirect from old URLs to new on transfer; existing forks, clones, and `origin` remotes keep working.
  - **Phase 5a-2 &mdash; migrate the gitbook spaces.** Tracked in [#147](https://github.com/saveenr/VisioAutomation/issues/147) (milestone `2026-12`). Custom redirects (or placeholder pages) cover the old URLs.
  - **Phase 5b &mdash; in-repo URL rewrite.** *(single mechanical commit, after 5a-1 and 5a-2 are done)*
    - Sed-replace `github.com/saveenr` &rarr; new owner, and `saveenr.gitbook.io` &rarr; new gitbook subdomain, across the inventory below. Skip past CHANGELOG entries (frozen historical record) and git history.
  - **Phase 5c &mdash; leave redirects in place permanently.** GitHub's auto-redirect doesn't expire and gitbook's custom redirects persist.
- **Inventory** of files containing `saveenr` URLs (for Phase 5b):
  - `readme.md` (badge URL, gitbook user-guide links, copyright line).
  - `CLAUDE.md`, `docs/ROADMAP.md`, `docs/COMPLETED.md`, `docs/OVERVIEW.md`, `docs/internal/custom-property-encoding.md`, `docs/futures/*.md` (links to issues / PRs / gitbook pages, both in body and in cross-refs).
  - `NuGet/CHANGELOG.md`, `VisioAutomation_2010/VisioPowerShell/CHANGELOG.md` (issue links from prior entries &mdash; **prefer to leave these as-is** since they're frozen historical record; the redirects ensure they keep working).
  - `VisioAutomation_2010/VisioPowerShell/Visio.psd1` (`PrivateData` URLs &mdash; ProjectUri, LicenseUri, etc., if present).
  - ~~`VisioAutomation_2010/VTest/datafiles/directed_graph_1.xml` (XML schema-reference comment pointing at gitbook).~~ Already handled 2026-05-07: the comment was throwaway-informational, so removed entirely rather than rewritten. One less file for Phase 5b.
  - `.github/workflows/release-nuget.yml` and `.github/workflows/release-psmodule.yml` (release-notes templates linking back at github.com/saveenr/... and saveenr.gitbook.io/...).
- **Decision points still open** (the two issues both note these):
  - Destination GitHub account/org name (e.g., `sevenpens`, `SevenPens`, or some other org).
  - Destination gitbook account/subdomain.
  - Whether the two moves happen at the same time, or stagger.
- **Effort:** Phase 5a is user-side &mdash; ~30 min for the GitHub transfer, ~15 min per gitbook space migration. Phase 5b is M, ~13 files of mechanical replacement; one-shot commit after 5a verifies.

#### Axis 6: Code-comment references *(done 2026-05-07)*
- [`GridOfMasters.cs:12`](../../VisioAutomation_2010/VSamples/Samples/Misc/GridOfMasters.cs): the dead-MSDN-blog comment line removed entirely. The URL pointed at an old `blogs.msdn.com/saveenr/...` post; the MSDN blogs platform was retired, so the URL was already dead.
- [`DemoIronPython/demo.py:109`](../../VisioAutomation_2010/DemoIronPython/demo.py): `r"D:\saveenr\data1.csv"` &rarr; `r"D:\sample\data1.csv"` (neutral placeholder path).

#### Axis 7: Legal copyright in LICENSE.txt *(done 2026-05-07; brand swap)*
- **Decision recorded 2026-05-07:** treat the change as a brand swap. SevenPens is the handle / pen-name the same legal person uses; legal authorship of the code traces through the historical record (git author lines, the LICENSE file in earlier tags, etc.) without depending on the current LICENSE.txt's exact spelling.
- Applied:
  - [`LICENSE.txt`](../../LICENSE.txt) line 3: `Copyright (c) 2016 Saveen Reddy` &rarr; `Copyright (c) 2016 SevenPens`. Year preserved.
  - [`readme.md`](../../readme.md) license line: `[MIT](LICENSE.txt). Copyright (c) Saveen Reddy.` &rarr; `[MIT](LICENSE.txt). Copyright (c) SevenPens.`
- If the situation ever changes (e.g., SevenPens becomes a real legal entity that owns the IP via formal assignment), the LICENSE.txt line should be re-revisited then. For now the displayed-author and legal-copyright stories are coherent at "SevenPens" across all surfaces.

#### Axis 8: Test fixtures *(done 2026-05-07)*
- The XML / log fixtures under [`VisioAutomation_2010/VTest/datafiles/`](../../VisioAutomation_2010/VTest/datafiles/) (`XMLErrorLog_Visio_2010_1.txt`, `XMLErrorLog_Visio_2013_1.txt`, `VSDX_Log_Visio_2013.txt`) used to embed machine paths and hostnames from the user's 2015 dev machines (`C:\Users\Saveen\`, `C:\Users\saveenr\`, `Saveen_ASGARD9`, `saveenr_SAVEENR3`). All three files have been mechanically scrubbed:
  - `Saveen_ASGARD9` &rarr; `Tester_TESTBOX`
  - `saveenr_SAVEENR3` &rarr; `tester_TESTBOX`
  - `Saveen` &rarr; `Tester` (in path components like `C:\Users\Saveen\...`)
  - `saveenr` &rarr; `tester` (in path components like `C:\Users\saveenr\...`)
  - `SAVEENR` &rarr; `TESTBOX` (in any leftover hostname suffixes)
- **Verification:** [`XmlErrorLogTests.cs`](../../VisioAutomation_2010/VTest/Core/Application/XmlErrorLogTests.cs) confirms its assertions only use `EndsWith` on filenames plus session/record counts and types &mdash; none check the user/machine-name substrings. The three relevant tests (`VSD_Load_Visio2013`, `XmlErrorLog_Load_Visio2010_1`, `XmlErrorLog_Load_Visio2013_1`) all pass post-scrub. They're pure file-I/O parser tests and run in under 1s total without needing a live Visio.
- **Out of scope of this axis:** historical CHANGELOG entries, prior commit messages, and the git author lines themselves still reference saveenr where they did at write time. Rewriting the past has a much bigger blast radius (commit hashes change, every external link to a specific commit breaks, tag verification fails) and gains nothing branding-wise. Don't.

#### Axis 9: Retire unused legacy publish-identity accounts *(tracked in [#148](https://github.com/saveenr/VisioAutomation/issues/148))*
- **What:** the [`Visio` PSGallery module](https://www.powershellgallery.com/packages/Visio) currently has three owners: `saveenr` (historical, retained), `SevenPens` (active going forward), and `VisioAutomation` (created years ago to separate work-from-hobby identity, never actually used). The `VisioAutomation` co-owner adds surface area without value and should be retired.
- **Why this is its own axis:** discovered 2026-05-07 while confirming SevenPens co-ownership. Wasn't part of the original 8-axis plan because the account hadn't surfaced. Now that `SevenPens` actively fulfills the "generic dev / team identity" role, the `VisioAutomation` account is redundant.
- **Plan:** remove `VisioAutomation` from the PSGallery owners list; delete the underlying account if PSGallery supports it, otherwise leave dormant. Check nuget.org and other feeds for the same identity for symmetry. Detail in [#148](https://github.com/saveenr/VisioAutomation/issues/148).
- **Effort:** S, user-side, no repo changes. Independent of other axes (in particular, doesn't gate or get gated by Axis 5).

#### Cross-refs

- [`releases.md`](releases.md#microsoft-package-compliance-gate-on-the-saveenr-nugetorg-account-operational-quirk-discovered-2026-05-07-during-the-300-publish) for the operational quirk that drove Axis 2.
- The [`nuget_publish_identity.md` project memory](../../../../.claude/projects/C--Users-savee-Documents-GitHub-VisioAutomation/memory/nuget_publish_identity.md) for the sticky operational rule on Axes 2 and 3 (the rule covers both feeds).
- [`docs.md`](docs.md) for the gitbook-side identity question, which couples to Axis 5.
- GitHub issues: [#146](https://github.com/saveenr/VisioAutomation/issues/146) (Axis 5a-1, GitHub repo move), [#147](https://github.com/saveenr/VisioAutomation/issues/147) (Axis 5a-2, gitbook moves), [#148](https://github.com/saveenr/VisioAutomation/issues/148) (Axis 9, retire VisioAutomation account). Issues 146 and 147 are scheduled to the [`2026-12` milestone](https://github.com/saveenr/VisioAutomation/milestone/1).
