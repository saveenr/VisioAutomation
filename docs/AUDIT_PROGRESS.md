# User-facing docs audit — progress

Working file for the Phase 1 doc-audit work. Tracks which pages have been audited, what was found, and what still needs to be done. **Delete this file when the audit is complete and the fixes have been merged into the gitbook repos.**

## Scope

Two gitbook source repos, cloned as siblings of `VisioAutomation/`:

| Repo | Branch audited | .md files | Lines | Last meaningful commit |
|---|---|---|---|---|
| [`VisioAutomation_GitBook_Docs`](https://github.com/saveenr/VisioAutomation_GitBook_Docs) | `main` | 17 | 684 | 2026-05-03 (squashed gitbook syncs) |
| [`VisioPowerShellDocs`](https://github.com/saveenr/VisioPowerShellDocs) | **`visiops_v4_docs`** | 116 | 2,214 | 2026-05-03 (`ccaf78d`, recent gitbook sync) |

The PS docs repo also has a `visiops_v3_docs` branch (older). The current PS module ships at v4.6.0, so the v4 doc branch is the right canonical target — it just trails the module by 0.6 minor versions.

## Status legend

- ⬜ not started
- 🟡 in progress
- ✅ audited — no fixes needed
- 🔧 audited — fixes identified (see Findings below)

## VisioAutomation_GitBook_Docs (`main` branch)

| Page | Status |
|---|---|
| `README.md` | 🔧 |
| `SUMMARY.md` | ⬜ |
| `compiling.md` | ⬜ |
| `stencils-and-masters.md` | ⬜ |
| `extension-methods.md` | ⬜ |
| `shapesheet/README.md` | ⬜ |
| `shapesheet/cells.md` | ⬜ |
| `shapesheet/query-the-shapesheet.md` | ⬜ |
| `shapesheet/modify-the-shapesheet.md` | ⬜ |
| `user-defined-cells.md` | ⬜ |
| `convert-values.md` | ⬜ |
| `custom-properties.md` | ⬜ |
| `resources/README.md` | ⬜ |
| `resources/stencils-and-masters.md` | ⬜ |
| `namespaces.md` | ⬜ |
| `classes.md` | ⬜ |
| `related-projects.md` | ⬜ |

## VisioPowerShellDocs (`visiops_v4_docs` branch)

Structural pages:

| Page | Status |
|---|---|
| `README.md` | 🔧 |
| `SUMMARY.md` | 🔧 |
| `links.md` | 🔧 |
| `quick-start.md` | 🔧 |

Section landings + content:

| Section | Files | Lines | Status |
|---|---|---|---|
| `basics/` | 15 | 457 | ⬜ |
| `cmdlets/` | 66 | 977 | ⬜ |
| `automatic-diagrams/` | 8 | 260 | ⬜ |
| `samples/` | 5 | 198 | ⬜ |
| `installation/` | 4 | 61 | ⬜ |
| `developer-info/` | 3 | 93 | ⬜ |
| `technical-notes/` | 11 | 168 | ⬜ |

---

## Findings

### VisioAutomation_GitBook_Docs/README.md 🔧

- Links to **`https://github.com/saveenr/VisioPowerShell/wiki`** for the PowerShell module — that separate repo doesn't exist anymore; the PS module lives **inside** the main `VisioAutomation` repo. Replace with link to either the PS gitbook (`https://saveenr.gitbook.io/visiopowershell/`) or the in-repo `VisioAutomation_2010/VisioPowerShell/README.md`.
- Links to **`https://github.com/saveenr/Visio-Power-Tools/releases`** "no longer maintained" — verify the repo still exists, otherwise remove the line.
- VisioAutomation 2007 reference is fine to keep as historical context.

### VisioPowerShellDocs/README.md 🔧

- Front matter: `description: This documentation covers Visio PowerShell version 4.5.0` — current module version on the PowerShell Gallery is **4.6.0**.
- Mentions Visio 2010 compatibility (matches the module's own description). OK.
- "Alternatives" section recommends `VisioBot3000` — verify the repo is still maintained before keeping the recommendation.

### VisioPowerShellDocs/SUMMARY.md 🔧

- **Definitely wrong link** (line 100): `Set-VisioUserDefinedCell` points at `cmdlets/user-defined-cells/remove-visiocustomproperty.md` — wrong file (it's a *remove* cmdlet under the *custom-properties* topic). The Set-VisioUserDefinedCell page should presumably exist as its own file.
- **Definitely wrong link** (line 15): `Close Visio applications` (in `basics/`) points at `basics/get-userdefinedcell.md` — wrong file.
- **Wrong cmdlet names** (current code uses `VisioUserDefinedCell`, not `UserDefinedCell`):
  - Line 99: `Get-UserDefinedCell` → should be `Get-VisioUserDefinedCell`
  - Line 101: `Remove-UserDefinedCell` → should be `Remove-VisioUserDefinedCell`
- **Typo** (line 122): `[Publis to PowerShell Gallery]` — should be "Publish".
- `Select-VisioShape Invert` and `Select-VisioShape -None` (lines 89-90) are listed as separate cmdlets but they are parameter variations of `Select-VisioShape`. Consider folding into one entry or being explicit they're parameter usages.
- **3 `[TBD]` markers** indicate pages that were never written: `Copy-VisioPage [TBD]` (line 71), `Select-VisioPage [TBD]` (line 78), `Get-VisioText [TBD]` (line 106). Note: `Copy-VisioPage` *does* exist in the current code (`Commands/VisioPage/CopyVisioPage.cs`), so the page can be written. Same for `Get-VisioText` (`Commands/VisioText/GetVisioText.cs`). `Select-VisioPage` also exists. All three TBDs can be filled in.
- **Cmdlet inventory delta:** see *Cmdlet Inventory* section below.

### VisioPowerShellDocs/links.md 🔧

- TechNet blog link `http://blogs.technet.com/b/heyscriptingguy/...` — Microsoft retired the `blogs.technet.com` domain years ago. Likely dead / 404. Replace with the migrated URL on `learn.microsoft.com` or remove.
- Third-party tools `VisioBot3000` and `PSVA` — verify the repos still exist and are not abandoned before keeping the links.

### VisioPowerShellDocs/quick-start.md 🔧

- **Line 17:** `Open-VisioDocument "basic_u.vss"` — the legacy `.vss` stencil format may still work but `.vssx` is modern. Verify the example still loads on a current Visio.
- **Code blocks have no language tag** (lines 5-7, 11-23) — adding `powershell` enables syntax highlighting on the gitbook page.
- **Line 33:** prose says *"The variable `$p` is defined …"* but the script uses `$points`. Should say `$points`.
- **Asset reference** `.gitbook/assets/snap00001.png` — verify the file exists in the repo (it's in the gitbook-managed folder, may or may not be tracked).

---

## Cmdlet inventory (code v4.6.0 vs docs `visiops_v4_docs`)

64 cmdlets in current code; 47 distinct verb-noun cmdlet links in SUMMARY. Spot-checked the section landing pages (`pagecells.md`, `shapecells/`, `selection/`, `container.md`) for hidden coverage of "missing" cmdlets.

### Wrong cmdlet names in docs (2) 🔧
Both already noted under SUMMARY findings — bear repeating here as inventory issues:
- `Get-UserDefinedCell` → should be `Get-VisioUserDefinedCell`
- `Remove-UserDefinedCell` → should be `Remove-VisioUserDefinedCell`

### Documented cmdlets that don't exist in code (2) 🔧
Confirmed via grep across the entire `VisioPowerShell` project — neither name appears anywhere in code. Either removed or never implemented:
- `Export-VisioSelection` (doc page: `cmdlets/selection/export-the-current-selection.md`)
- `Test-VisioSelectedShapes` (doc page: `cmdlets/selection/checking-selection-status.md`)

Resolution choice: remove the doc pages, **or** investigate whether equivalent functionality exists under different cmdlets (e.g., `Get-VisioShape` for "what's selected", `Export-VisioShape` against the selection).

### Cmdlets in code but undocumented (~17, after accounting for section-page coverage) 🔧

Sorted into rough categories based on user-visibility:

**Central cmdlets that should have docs:**
- `New-VisioShape` — primary shape-creation cmdlet, not documented (!)
- `Remove-VisioShape`
- `Set-VisioPageCells`
- `New-VisioPageCells`

**Cmdlet families with partial doc coverage but missing pages:**
- *Control* family (`Get-VisioControl`, `New-VisioControl`, `Remove-VisioControl`) — no `controls/` section in docs at all
- *PageCells* — single `cmdlets/pagecells.md` exists but is buggy (see Section page bugs below)
- *ShapeCells* — has a section but no per-cmdlet pages for `Get-VisioShapeCells`, `New-VisioShapeCells`
- *Container* — has an essentially empty page; `New-VisioContainer` undocumented

**Smaller / utility cmdlets** (debatable whether worth documenting):
- `Get-VisioClient`, `Get-VisioLockCells`, `Import-VisioModel`, `Measure-VisioShape`, `New-VisioPoint`, `New-VisioRectangle`, `Select-VisioDocument`, `Test-VisioDocument`

### Section page bugs 🔧

- **`cmdlets/pagecells.md`** lists *Shape* cmdlets where it should list *Page* cmdlets:
  ```
  These cmdlets work with the ShapeSheet of pages:
      New-VisioShapeCells   ← should be New-VisioPageCells
      Get-VisioPageCells
      Set-VisioShapeCells   ← should be Set-VisioPageCells
  ```
  Two of three names are wrong.
- **`cmdlets/container.md`** is essentially a placeholder — just the `# Container` heading. Needs content for `New-VisioContainer` at minimum.

### Completeness target for Phase 2: **Option 2 — Reasonable completeness** ✓

Decided 2026-05-03. Fix what's wrong (strict accuracy) **and** write new doc pages for the central undocumented cmdlets. Don't try to document the small/utility cmdlets — they get a single "known undocumented cmdlets" note in the docs and that's it.

---

## Phase 2 doc-fix work plan

Concrete list of what needs to land before the Phase 2 release ships, derived from the findings above.

### A. Strict-accuracy fixes (PS docs) ✅ done

All landed as **local commits** on `visiops_v4_docs` in `VisioPowerShellDocs`. **Not yet pushed.**

| # | Item | Status | Commit |
|---|---|---|---|
| 1 | `README.md` version pin 4.5.0 → 4.6.0 | ✅ | `e3ba7cf` |
| 2a | SUMMARY line 15 — wrong link target for "Close Visio applications" | ✅ via rename | `4677f63` |
| 2b | SUMMARY lines 99-101 — UserDefinedCell wrong names + wrong link target | ✅ | `5e8624d` + heading-fix follow-up `f66ce8d` |
| 2c | SUMMARY line 122 — `Publis` → `Publish` typo | ✅ | `0bc600c` |
| 2d | SUMMARY lines 89-90 — Select-VisioShape variations | **deferred** — not strictly wrong, just stylistic; revisit if time |
| 2e | SUMMARY lines 71, 78, 106 — `[TBD]` markers | **deferred to Section C** — link targets already correct; markers honestly flag incomplete content (Copy-VisioPage, Select-VisioPage, Get-VisioText pages need to be written) |
| 3 | `links.md` — dead TechNet link | ✅ removed | `1798a0b` |
| 4 | `quick-start.md` — `$p`/`$points`, code-block tags, grammar | ✅ | `8186f71` |
| 4a | `quick-start.md` — verify `basic_u.vss` still loads on current Visio | **deferred — needs live Visio test** |
| 5 | `cmdlets/pagecells.md` content bug | ✅ | `3cd1f29` |
| 6+7 | `cmdlets/selection/` — both pages document non-existent cmdlets | ✅ entire empty section deleted (3 files) | `e1d5762` |

### B. Strict-accuracy fixes (.NET docs)

In `VisioAutomation_GitBook_Docs` on `main`. **Pushed.**

| # | Item | Status | Commit(s) |
|---|---|---|---|
| 8 | `README.md` — stale `VisioPowerShell/wiki` link + intro grammar | ✅ | `c1ddd25` |
| 8a | `compiling.md` — VS 2019 → 2022, C# 7.0 → 8.0, link to in-repo BUILDING | ✅ | `1b6e4b6` |
| 8b | `convert-values.md` — entire page deleted (documents non-existent `VisioAutomation.Convert` static class) | ✅ | `b95b381` |
| 8c | `user-defined-cells.md` — `UserDefinedCellsHelper` → `UserDefinedCellHelper` (singular) | ✅ partial | `83ae5a4` |
| 8d | `custom-properties.md` — wrong namespaces (`Shapes.CustomProperties.X`) + truncated last line + duplicate closing code fence | ✅ partial | `a278563` + `f6c202b` |
| 8e | `stencils-and-masters.md` + `extension-methods.md` — replace dead `VisioAutomation.Drawing.{Point,Size}` with `VisioAutomation.Core.{Point,Size}`; fix broken `=>` lambda formatting | ✅ | `6adfb83` |
| 8f | `shapesheet/cells.md` — systemic API drift fix: `SRC` → `Src`, `VA.ShapeSheet.SRCConstants` → `VA.Core.SrcConstants`, constant names dropped underscores (`Char_Color` → `CharColor`, `FillForegnd` → `FillForeground`, etc.); deleted whole "Converting between cell names and (s,r,c)" section (documented non-existent `ShapeSheetHelper.GetSRCFromName`) | ✅ | `92104df` |
| 8g | `shapesheet/query-the-shapesheet.md` — wrong API (`AddColumn` → `Columns.Add`), wrong type (`int[] results_bool` → `bool[]`), syntax bugs (`var query.AddColumn`, `new int { … }`), typos (`intm`, `twoc olumns`, `perfoming`, `Retieving`) | ✅ partial | `3e5b66e` |
| 8h | `shapesheet/modify-the-shapesheet.md` — rewrote for current API: `SRCUpdate`/`SIDSRCUpdate` → `SrcWriter`/`SidSrcWriter`, `SetFormula`/`SetResult`/`Execute` → `SetValue`/`Commit(target, CellValueType)`, fix `shape.Cells["Pinx"]` → `shape.CellsU["PinX"]` | ✅ | `c1b800a` |
| 9 | `SUMMARY.md` link-target audit — programmatically verified all `(file.md)` targets resolve | ✅ | (no commit needed — clean) |
| 11a | `shapesheet/README.md` — wrote a real section landing intro | ✅ | `cfc94fb` |
| 11b | `classes.md` / `namespaces.md` / `related-projects.md` — added context paragraph above each diagram, expanded related-projects with VisioPS + PSVA. Diagram-currency still needs visual verification by user. | ✅ | `e3a514c` |
| 11c | `resources/stencils-and-masters.md` — duplicate page deleted; SUMMARY entry removed | ✅ | `641083a` |

**Section B done.** All accuracy fixes for the .NET docs are pushed.

**Section B follow-ups (originally Section C carry-overs) — done:**
- ✅ `user-defined-cells.md` rewritten using `GetDictionary(shape, CellValueType)`, `Set(shape, name, UserDefinedCellCells)`, `ShapeIDPairs.FromShapes(...)` for multi-shape, plus the `IsValidName`/`CheckValidName` methods that previously weren't mentioned. Commit `08d4e9f`.
- ✅ `custom-properties.md` rewritten with the full `CustomPropertyCells` field set (Value, Prompt, Label, Format, Type, Calendar, Invisible, LangID, SortKey, Ask), correct `GetDictionary` calls, and the `ShapeIDPairs` multi-shape pattern. Commit `be43089`.
- ✅ `query-the-shapesheet.md` rewritten with a return-type table making clear which shape each method returns (`DataRows<T>` / `DataRowGroup<T>` / `DataRowGroups<T>`); fixed indexer syntax (`[r][c]` not `[r,c]`); fixed multi-shape SectionQuery to use `Core.ShapeIDPairs` instead of `IList<int>`; restructured the grouping discussion. Commit `e0f50f7`.

### C. New doc pages to write (PS docs, option 2 additions)

All landed as **local commits** on `visiops_v4_docs` in `VisioPowerShellDocs`. **Not yet pushed.**

| # | Item | Status | Commit |
|---|---|---|---|
| 12 | `cmdlets/shapes/new-visioshape.md` — central cmdlet, all six parameter sets (drop master, rectangle, oval, line, polyline, bezier) plus `-Cells`. Drive-by fixes to existing pages whose `New-VisioShape` examples wouldn't actually run (`connect-shapes.md`, `examples.md`, `basics/drop-masters.md`). | ✅ | `e554d3c` |
| 13 | `cmdlets/shapes/remove-visioshape.md` | ✅ | `ca8b322` |
| 14 | `cmdlets/pages/new-visiopagecells.md`, `set-visiopagecells.md` — new cmdlet pages. Fixed bugs in `cmdlets/pagecells.md` landing (`-Pages` → `-Page`, `$pages`/`$page` mix-up, code-block tags). SUMMARY entries nested under PageCells (not Pages, to avoid confusing them with page-lifecycle cmdlets). | ✅ | `bbd9241` |
| 15 | `cmdlets/shapecells/new-visioshapecells.md`, `get-visioshapecells.md` — new cmdlet pages. Replaced the empty `shapecells/README.md` with a real landing. Drive-by fixes to `working-with-shape-cells.md` and `format-text.md`: `-Shapes` → `-Shape`, broken `New-VisioShape $master 2,2` form, code-block tags. | ✅ | `5e881c5` |
| 16 | `cmdlets/control/` — entirely new section (4 files). README explains control handles as ShapeSheet rows. SUMMARY entry slotted alphabetically between Container and Custom properties. | ✅ | `8b098b1` |
| 17 | `cmdlets/container.md` — was a one-line placeholder. Now defines what a Visio container is, documents `New-VisioContainer`, and shows the select-then-drop pattern. | ✅ | `7746f6b` |
| 18 | New `cmdlets/other-cmdlets.md` lists all eight small/utility cmdlets with one-line descriptions and points at `Get-Help` for full reference. SUMMARY entry added at the bottom of the Cmdlets section. Drive-by fix to `basics/list-of-all-cmdlets.md`: removed `Get-VisioLayer` (not in code), added the missing `New-VisioPoint` and `New-VisioRectangle` entries. | ✅ | `4f763db` |

**Section C done.** All seven new-page items landed; pushed to `origin/visiops_v4_docs`.

### D. Stub-fill pass (post-Section-C, doc-only)

A second pass to fill the bare-headline stub pages flagged at the end of Section C. Doc-only — code fixes for the bugs surfaced below are deferred pending a release-process discussion.

| Page | Status | Notes |
|---|---|---|
| `cmdlets/shapes/copy-visioshape.md` | ✅ | |
| `cmdlets/shapes/test-visioshape.md` | ✅ | |
| `cmdlets/pages/measure-visiopage.md` | ✅ | |
| `cmdlets/shapes/export-visioshape.md` | ✅ | Initially documented `-Overwrite` workaround; rewritten in Phase A (after 4.6.1 fix) to describe normal behavior. |
| `cmdlets/shapes/lock-visioshape.md` | ✅ | Written for 4.6.1 release once the binder bug was fixed. |
| `cmdlets/shapes/unlock-visioshape.md` | ✅ | Same as Lock. |

**Stubs found during this pass that were NOT in Section D scope** (mostly section READMEs and the entire `cmdlets/visioapplication/` per-cmdlet folder; also `cmdlets/text/get-visiotext.md` and `cmdlets/pages/select-visiopage-tbd.md`). Tracked here for visibility but no decision yet on whether to fill them.

### E. Bare-headline stub fill + template overlay (post-4.6.1 release)

A multi-stage pass after the 4.6.1 release shipped, applying a consistent **Syntax + Parameters** documentation template across every cmdlet page and filling the remaining bare-headline stubs.

**Phase A — Template overlay across content-rich pages.** ~47 cmdlet pages updated, organized into nine section commits. The template adds a `## Syntax` block (PowerShell `Get-Help -Syntax` style; one block per parameter set for multi-set cmdlets), a `## Parameters` table, and standardizes the `## Examples` and `## See also` sections. Bold for cmdlet name in topic-establishing prose; backticks for inline code references. Section commits on `visiops_v4_docs`:

| Section | Commit | Pages | Notes |
|---|---|---|---|
| Prototype + Documents | `0495191`, `b77c06a` | 6 | `Close-VisioDocuments` typo, `Set-VisioDocument` → `Select-VisioDocument`, "if not needed" wording |
| Pages | `92f72ef` | 8 | `Set-VisioPage` (not a cmdlet), `Get-VisioScriptingClient` (not a cmdlet), `-AllPages` (not a parameter), positional binding on Remove-VisioPage |
| Shapes | `fb2fb18` | 15 | `ShapeSelectionOperation` enum value mismatch (`All`/`None`/`Invert` → `SelectAll`/`SelectNone`/`InvertSelection`), `New-VisioGroup` (not a cmdlet), `-Operation` → `-SelectionOperation`, `-Recursive` (not real), `-NudgeX` for vertical nudges |
| ShapeCells | `8a15b23` | 2 | |
| Hyperlinks | `83e5689` | 3 | Many missing parameters surfaced |
| Custom-properties | `8aed4be` | 4 | `-Master`/`-Stencil` on Get-VisioMaster (not real), `-Shapes` plural variants |
| User-defined-cells | `6b5df9a` | 3 | |
| Text/Master/Control/Windows | `0eaac0e` | 6 | `VisioAutomation.Geometry.Point` (wrong namespace), Format-VisioWindow prose param names |

**Phase B — Bare-headline stubs filled.** All previously empty cmdlet pages now have content with the Phase A template applied:

| Page | Commit |
|---|---|
| `cmdlets/visioapplication/{close,get,new,out,test,undo,redo}-visioapplication.md` (7 pages) | `fc06925` |
| `cmdlets/pages/invoke-visioduplicate-page.md` (Copy-VisioPage, was [TBD]) | `fc06925` |
| `cmdlets/pages/select-visiopage-tbd.md` (Select-VisioPage, was [TBD]) | `fc06925` |
| `cmdlets/text/get-visiotext.md` (was [TBD]) | `fc06925` |

SUMMARY: dropped the three `[TBD]` markers since the pages now have content.

**Phase C — Multi-page Select-VisioShape consolidation.** Folded into Phase A's Shapes commit (`fb2fb18`). The three pages (`selecting-shapes.md`, `invert-the-selection.md`, `clearing-the-selection.md`) now share consistent enum values; the two variant pages each link back to the canonical page.

### Code-level findings (resolved in 4.6.1)

The three code-level findings flagged during Sections C and D were all fixed and shipped in the **Visio PowerShell 4.6.1** release on 2026-05-03 (tag `VisioPS_4.6.1` on `2026_Refresh`):

1. ✅ **`Lock-VisioShape` / `Unlock-VisioShape` switches now bind.** Added `[SMA.Parameter(Mandatory = false)]` to all 20 switch fields on both classes.
2. ✅ **`Export-VisioShape` file-exists check.** Inverted `!File.Exists` flipped to `File.Exists`.
3. ✅ **`NewVisioShape._check_num_Points` guards.** `throw` added before each constructed `ArgumentOutOfRangeException`.

The Phase A doc rollout describes the post-4.6.1 behavior; the published `lock-visioshape.md` / `unlock-visioshape.md` pages also include a "Requires 4.6.1 or later" note.

### Workflow questions

- **Branching.** Commit doc fixes directly to `visiops_v4_docs` / `main`, or use feature branches and PRs? My instinct: commit directly — small project, no review chain — but easy to switch later.
- **Authorship.** When committing in the gitbook repos, what author identity should the commits carry? My current git config is the same as the main repo (`TheSevenPens`). If you want a different identity for these commits (e.g., your name), I can override per-commit with `--author`.

---

## Open questions for later

- **Where the docs should ultimately live.** The Phase 3 FUTURES item *"Decide where docs live long-term"* — the audit might surface enough drift / stale infrastructure that consolidation becomes more attractive than fixing in place. Track and revisit.
- **Branch versioning policy:** decided. **Only audit/fix `visiops_v4_docs`.** `visiops_v3_docs` is a frozen historical record for users still on the v3 module. No cherry-picking; no syncing. Right answer for v3 users hitting issues is "upgrade to v4."

---

## Next steps

1. ✅ Initial scoping
2. ✅ Structural-page audit (READMEs, SUMMARYs, links, quick-start)
3. ⬜ Audit `SUMMARY.md` of VisioAutomation gitbook
4. ⬜ Per-cmdlet inventory diff (66 docs entries vs 64 code files): identify renames, removals, additions
5. ⬜ Audit each section folder of the PS docs in turn (`basics/`, `cmdlets/`, …)
6. ⬜ Audit each substantive page of `VisioAutomation_GitBook_Docs` (`shapesheet/*`, `convert-values.md`, etc.)
7. ⬜ Once findings are complete, propose fix PRs against each gitbook repo
