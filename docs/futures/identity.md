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

#### Axis 3: PSGallery publishing identity *(pending; small effort)*
- **What:** mirror the nuget.org workaround on PSGallery. Have `SevenPens` request co-owner status on the [`Visio` module](https://www.powershellgallery.com/packages/Visio) (saveenr the current sole owner), then generate a SevenPens-scoped API key and rotate the GitHub repo's `PSGALLERY_API_KEY` secret to that key.
- **Why:** symmetry with nuget.org. Also future-proofs against PSGallery enforcing analogous Microsoft-package compliance rules &mdash; the same story would play out and the workaround would be the same. Doing it preemptively means we never get caught with a failed publish.
- **How:** PSGallery owner-management is at https://www.powershellgallery.com/packages/Visio/Manage. Co-owner invitation flow same shape as nuget.org. After secret rotation, `publish-psmodule.yml` runs unchanged.
- **Cost of waiting:** zero today (the saveenr key still works on PSGallery), but if PSGallery tightens enforcement we hit the same surprise that happened with nuget.org 3.0.0 on 2026-05-07.
- **Effort:** S (~30 min user-side; no repo changes).

#### Axis 4: Display authorship in artifact metadata *(pending; decision-coupled)*
- **What:** the displayed-author / copyright fields in shipped artifact metadata are still `Saveen Reddy` / `saveenr`:
  - [`VisioAutomation_2010/VisioPowerShell/Visio.psd1`](../../VisioAutomation_2010/VisioPowerShell/Visio.psd1) line 17: `Author = 'Saveen Reddy'`
  - [`VisioAutomation_2010/VisioPowerShell/Visio.psd1`](../../VisioAutomation_2010/VisioPowerShell/Visio.psd1) line 23: `Copyright = 'Saveen Reddy'`
  - [`NuGet/VisioAutomation2010.nuspec`](../../NuGet/VisioAutomation2010.nuspec) line 5: `<authors>saveenr</authors>`
  - [`NuGet/VisioAutomation2010.nuspec`](../../NuGet/VisioAutomation2010.nuspec) line 15: `<copyright>Copyright Saveen Reddy</copyright>`
  - [`VisioAutomation_2010/VPlayground/Properties/AssemblyInfo.cs`](../../VisioAutomation_2010/VPlayground/Properties/AssemblyInfo.cs) line 9: `[assembly: AssemblyCopyright("Copyright Saveen Reddy")]`
- **Why this is a decision, not just a rename:** the displayed-author field is *separate from legal copyright* (Axis 7 below). Many open-source projects display a brand or handle in `Author` while legal copyright stays in LICENSE.txt, but the two need to tell a coherent story. Three options:
  - **A &mdash; Author "SevenPens", Copyright "Saveen Reddy":** displayed author is the brand, legal copyright is the human. Common pattern for solo developers using a publishing handle.
  - **B &mdash; Both "SevenPens":** treats SevenPens as the canonical name in all artifact-facing surfaces. Legal copyright in LICENSE.txt then needs to align (Axis 7).
  - **C &mdash; Both "Saveen Reddy":** keep status quo on these specific fields; identity transition stops at publishing/commit identity.
- **Effort to apply once decided:** S. Five line-edits across three files. The .nupkg's nuspec is regenerated on every release, so it picks up the new value at the next NuGet release; the .psd1 ships as-is on the next PSGallery release.
- **Coupling:** to Axis 7. Don't change displayed-author without thinking about LICENSE.txt at the same time.

#### Axis 5: Hosting URLs (`github.com/saveenr/...`, `saveenr.gitbook.io/...`) *(pending; gated on external moves)*
- **What:** every link to a hosted resource currently points at saveenr-owned spaces:
  - **GitHub:** `https://github.com/saveenr/VisioAutomation/...` (issues, PRs, file links, badges).
  - **gitbook:** `https://saveenr.gitbook.io/visioautomation/`, `https://saveenr.gitbook.io/visiopowershell/`.
- **Why URL replacement isn't a string-find-and-replace:** the URLs follow the underlying resource location. To change `github.com/saveenr/VisioAutomation` &rarr; `github.com/sevenpens/VisioAutomation` we have to actually transfer or fork the GitHub repo first. Same for the gitbook spaces. Until those moves happen, replacing the URLs in code/docs would just produce 404s.
- **Inventory** of files containing saveenr URLs (for the eventual rewrite pass &mdash; do not edit before the underlying moves):
  - `readme.md` (badge URL, gitbook user-guide links, copyright line).
  - `CLAUDE.md`, `docs/ROADMAP.md`, `docs/COMPLETED.md`, `docs/OVERVIEW.md`, `docs/internal/custom-property-encoding.md`, `docs/futures/*.md` (links to issues / PRs / gitbook pages, both in body and in cross-refs).
  - `NuGet/CHANGELOG.md`, `VisioAutomation_2010/VisioPowerShell/CHANGELOG.md` (issue links from prior entries).
  - `VisioAutomation_2010/VisioPowerShell/Visio.psd1` (`PrivateData` URLs &mdash; ProjectUri, LicenseUri, etc., if present).
  - `VisioAutomation_2010/VTest/datafiles/directed_graph_1.xml` (XML schema-reference comment pointing at gitbook).
  - `.github/workflows/release-nuget.yml` and `.github/workflows/release-psmodule.yml` (release-notes templates linking back at github.com/saveenr/... and saveenr.gitbook.io/...).
- **Open question:** is the move actually planned? GitHub repo transfer is irreversible-ish (URL redirects work but are best-effort) and affects every external consumer who has the URL bookmarked. gitbook space migration is also externally visible. Until these moves are committed to, leave the URLs alone.
- **Effort to apply once moves happen:** M. ~13 files with URL-shape replacements; mostly mechanical sed-style rewrites once the destination URL is known. CHANGELOG entries from past releases probably stay as-is (they're frozen historical record).

#### Axis 6: Code-comment references *(pending; trivial)*
- **What:** two source-file references that mention `saveenr` outside any author/copyright context:
  - [`VisioAutomation_2010/VSamples/Samples/Misc/GridOfMasters.cs:12`](../../VisioAutomation_2010/VSamples/Samples/Misc/GridOfMasters.cs): `// http://blogs.msdn.com/saveenr/archive/2008/08/06/visioautoext-...` &mdash; a comment pointing at an old MSDN blog post. **The MSDN blogs platform was retired**, so this URL is dead regardless. Could be removed entirely, or replaced with a working equivalent if the post was preserved on a successor blog.
  - [`VisioAutomation_2010/DemoIronPython/demo.py:109`](../../VisioAutomation_2010/DemoIronPython/demo.py): `>>> datatable = vi.Data.ImportCSV( r"D:\saveenr\data1.csv" )` &mdash; a fake example path in a docstring. Trivial rename to anything (`r"D:\sample\data1.csv"` etc.).
- **Effort:** S. Two one-line edits, no behavior change, no review subtlety.

#### Axis 7: Legal copyright in LICENSE.txt *(decision-needed; do not touch casually)*
- **What:** [`LICENSE.txt`](../../LICENSE.txt) line 3: `Copyright (c) 2016 Saveen Reddy`. [`readme.md`](../../readme.md) line 68 mirrors: `[MIT](LICENSE.txt). Copyright (c) Saveen Reddy.`
- **Why this is its own axis, not just another metadata rename:** LICENSE.txt names the legal copyright holder of the source code. It is not a display field. Changing it asserts that someone other than Saveen Reddy holds the copyright, which is only correct if (a) `SevenPens` is a legal entity (LLC, sole proprietorship, etc.) that has acquired the copyright by formal transfer, OR (b) `SevenPens` is being used as a pen-name for the same legal person. In case (b) the legal copyright is *still* held by Saveen Reddy regardless of what the file prints, and rewriting the LICENSE could create downstream legal ambiguity for consumers.
- **Decision options:**
  - **Keep `Copyright (c) 2016 Saveen Reddy` as-is.** Most defensible. The displayed brand can transition (Axis 4) without disturbing the legal record.
  - **Rewrite to a SevenPens entity** *only* if a real legal entity exists and there's an executed copyright assignment. Otherwise this is risky.
  - **Add a second line acknowledging the SevenPens brand** alongside the existing copyright line. Cosmetic, doesn't change the legal substance, makes the relationship visible to readers.
- **Coupling:** to Axis 4. The story across LICENSE.txt and the displayed-author fields should be coherent.
- **Effort:** S to apply, but the *decision* is the load-bearing part &mdash; not work that should be done without the user's explicit call.

#### Axis 8: Test fixtures and historical artifacts *(do not change)*
- The XML / log fixtures under [`VisioAutomation_2010/VTest/datafiles/`](../../VisioAutomation_2010/VTest/datafiles/) (`XMLErrorLog_Visio_2010_1.txt`, `XMLErrorLog_Visio_2013_1.txt`, `VSDX_Log_Visio_2013.txt`) contain captured 2015-era Visio error-log output that embeds machine paths like `C:\Users\Saveen\`, `C:\Users\saveenr\`, and machine names like `Saveen_ASGARD9` / `saveenr_SAVEENR3`. These are real recorded fixtures, used by the error-log parser tests, and tests likely depend on the exact bytes. Don't touch.
- Historical changelog entries, docs, and commit messages that reference saveenr likewise stay frozen &mdash; rewriting the past introduces drift between docs and the git history they reference.

#### Cross-refs

- [`releases.md`](releases.md#microsoft-package-compliance-gate-on-the-saveenr-nugetorg-account-operational-quirk-discovered-2026-05-07-during-the-300-publish) for the operational quirk that drove Axis 2.
- The [`nuget_publish_identity.md` project memory](../../../../.claude/projects/C--Users-savee-Documents-GitHub-VisioAutomation/memory/nuget_publish_identity.md) for the sticky operational rule on Axis 2 and the prereq for Axis 3.
- [`docs.md`](docs.md) for the gitbook-side identity question, which couples to Axis 5.
