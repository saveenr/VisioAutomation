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
| 9 | `SUMMARY.md` audit | ⬜ partial — incidental fixes in 8b deleted the `Convert values` entry; full link-target verification still pending |
| 11a | `shapesheet/README.md` — truly empty stub (just heading); fill or drop from SUMMARY | ⬜ |
| 11b | `classes.md` / `namespaces.md` / `related-projects.md` — diagram-only stubs; need diagram-currency verification by user (can't do from here) | ⬜ |
| 11c | `resources/stencils-and-masters.md` — smart quotes (Word-paste artifact), `// C# Visual Studio 2010` outdated comment, duplicates the root `stencils-and-masters.md` topic with different content (consolidation candidate) | ⬜ |

**Section B follow-ups deferred to Section C / option-2 work:**
- `user-defined-cells.md` and `custom-properties.md` both have method-signature drift: their `.Get(shape)` and `.Set(shape, name, value, prompt)` examples don't match the actual API, which uses `.GetDictionary(shape, CellValueType)` and `.Set(shape, name, CellPropertyCells)`. The class-name renames I did fix a third of the issue; full content rewrite needed for the rest.
- `query-the-shapesheet.md` describes return values as "Table object" but current code returns `Data.DataRows<T>` / `Data.DataRowGroup<T>` / `Data.DataRowGroups<T>`. Conceptual description still works at a high level but is technically inaccurate.

### C. New doc pages to write (PS docs, option 2 additions)

12. **`cmdlets/shapes/new-visioshape.md`** — central cmdlet; cover the full parameter set (`-Master`, `-Position`, plus the Rectangle / Oval / Line / Polyline / Bezier parameter sets).
13. **`cmdlets/shapes/remove-visioshape.md`** — pair with above.
14. **`cmdlets/pages/new-visiopagecells.md`**, **`set-visiopagecells.md`** — fix-up under the Pages section, plus updating the existing buggy `pagecells.md`.
15. **`cmdlets/shapecells/new-visioshapecells.md`**, **`get-visioshapecells.md`** — fill out the section.
16. **`cmdlets/control/`** — entirely new section. Add `README.md`, `get-visiocontrol.md`, `new-visiocontrol.md`, `remove-visiocontrol.md`. Add to SUMMARY.
17. **`cmdlets/container.md`** — currently a placeholder; flesh out for `New-VisioContainer`.
18. **One "Other cmdlets" / "Known undocumented cmdlets" note** — single page or an addendum to README listing the small/utility cmdlets we explicitly chose not to document (Get-VisioClient, Get-VisioLockCells, Import-VisioModel, Measure-VisioShape, New-VisioPoint, New-VisioRectangle, Select-VisioDocument, Test-VisioDocument).

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
