# User-facing docs audit — progress

Working file for the Phase 1 doc-audit work. Tracks which pages have been audited, what was found, and what still needs to be done. **Delete this file when the audit is complete and the fixes have been merged into the gitbook repos.**

## Scope

Two gitbook source repos, cloned as siblings of `VisioAutomation/`:

| Repo | Branch audited | .md files | Lines | Last meaningful commit |
|---|---|---|---|---|
| [`VisioAutomation_GitBook_Docs`](https://github.com/saveenr/VisioAutomation_GitBook_Docs) | `main` | 17 | 684 | 2026-05-03 (squashed gitbook syncs — content age unclear) |
| [`VisioPowerShellDocs`](https://github.com/saveenr/VisioPowerShellDocs) | `visiosps4.0.0docs` | 114 | 2,463 | 2019-10-13 |

The PS docs branch is pinned to module v4.0.0; the current module version is **4.6.0**, so the PS docs are roughly 6.5 years stale across 0.6 minor releases.

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

## VisioPowerShellDocs (`visiosps4.0.0docs` branch)

Structural pages:

| Page | Status |
|---|---|
| `README.md` | 🔧 |
| `SUMMARY.md` | 🔧 |
| `links.md` | 🔧 |
| `quick-start.md` | 🔧 |

Section landings + content (114 files total):

| Section | Files | Status |
|---|---|---|
| `basics/` | 12 | ⬜ |
| `cmdlets/` | 66 | ⬜ |
| `automatic-diagrams/` | 8 | ⬜ |
| `samples/` | 6 | ⬜ |
| `installation/` | 4 | ⬜ |
| `developer-info/` | 3 | ⬜ |
| `technical-notes/` | 11 | ⬜ |

---

## Findings

### VisioAutomation_GitBook_Docs/README.md 🔧
- Links to **`https://github.com/saveenr/VisioPowerShell/wiki`** for the PowerShell module — that separate repo doesn't exist anymore; the PS module lives **inside** the main `VisioAutomation` repo. Replace with link to either the PS gitbook (`https://saveenr.gitbook.io/visiopowershell/`) or the in-repo `VisioAutomation_2010/VisioPowerShell/README.md`.
- Links to **`https://github.com/saveenr/Visio-Power-Tools/releases`** "no longer maintained" — verify the repo still exists, otherwise remove the line.
- VisioAutomation 2007 reference is fine to keep as historical context.

### VisioPowerShellDocs/README.md 🔧
- Front matter: `description: This documentation is being updated to cover Visio PowerShell Version 4.4.0` — current module version is **4.6.0**.
- Vimeo demo link from ~2014; content acknowledged as older. Keep, since principles still apply.

### VisioPowerShellDocs/SUMMARY.md 🔧
- **Definitely wrong link:** `Set-VisioUserDefinedCell` points at `cmdlets/user-defined-cells/remove-visiocustomproperty.md` — wrong file (points to a *remove* cmdlet under the *custom-properties* topic).
- **Definitely wrong link:** `Close Visio applications` (in `basics/`) points at `basics/get-userdefinedcell.md` — wrong file.
- **Wrong cmdlet names** (current code uses `VisioUserDefinedCell`, not `UserDefinedCell`):
  - `Get-UserDefinedCell` → should be `Get-VisioUserDefinedCell`
  - `Remove-UserDefinedCell` → should be `Remove-VisioUserDefinedCell`
- `Select-VisioShape Invert` and `Select-VisioShape -None` are listed as separate cmdlets but they are parameter variations of `Select-VisioShape`. Consider folding into one entry or being explicit they're parameter usages.
- **12 `[TBD]` markers** in the table of contents — content was never written for those.
- **Cmdlet inventory delta:** 66 cmdlet entries here vs 64 cmdlet `.cs` files in current code. Need a per-cmdlet diff to identify (a) cmdlets renamed, (b) cmdlets removed, (c) cmdlets added since 2019. *(Deferred to per-cmdlet audit.)*

### VisioPowerShellDocs/links.md 🔧
- TechNet blog link `http://blogs.technet.com/b/heyscriptingguy/...` — Microsoft retired the `blogs.technet.com` domain years ago. Likely dead / 404. Replace with the migrated URL on `learn.microsoft.com` or remove.
- Third-party tools `VisioBot3000` and `PSVA` — verify the repos still exist and are not abandoned before keeping the links.

### VisioPowerShellDocs/quick-start.md 🔧
- **Line 30: stale namespace.** `New-Object VisioAutomation.Geometry.Point(4,5)` — that namespace does not exist in the current code. The current namespace for `Point` is `VisioAutomation.Core.Point` (per `docs/ARCHITECTURE.md`). The example as written will fail with a type-not-found error.
- **Line 28: `basic_u.vss`** — the legacy `.vss` stencil format may still work but `.vssx` is modern. Verify the example still loads on a current Visio.
- **Code blocks use `` ```text ``** instead of `` ```powershell `` (lines 7-11, 22-34). Switching gives proper syntax highlighting on the gitbook page.
- **Asset reference** `.gitbook/assets/snap00001.png` — verify the file exists in the repo (it's in the gitbook-managed folder, may or may not be tracked).

---

## Open questions for the next session

- **Should we keep both gitbook repos as the source of truth, or consolidate?** Phase 3 has a separate FUTURES item for "Decide where docs live long-term." The audit might surface enough drift that consolidation becomes more attractive — track findings here and revisit.
- **For broken example code in the docs (e.g., `VisioAutomation.Geometry.Point`),** do we fix in place against the current API, or fix against the API as it was when the doc was written (and bump the doc's version pin)? The PS docs are explicitly version-pinned to `visiosps4.0.0docs` — meaning the code examples there *should* match v4.0.0 of the module, not v4.6.0. If we fix against v4.6.0, we should rename the branch (`visiosps4.6.0docs`) or drop the version pin altogether.

---

## Next steps

1. ✅ Initial scoping (this commit)
2. ⬜ Finish auditing the structural pages of `VisioAutomation_GitBook_Docs` (`SUMMARY.md`)
3. ⬜ Per-cmdlet inventory diff: which docs cmdlets exist in current code, which don't, and vice versa
4. ⬜ Audit each section folder of the PS docs in turn (`basics/`, `cmdlets/`, …)
5. ⬜ Audit each substantive page of `VisioAutomation_GitBook_Docs` (`shapesheet/*`, `convert-values.md`, etc.)
6. ⬜ Once findings are complete, propose fix PRs against each gitbook repo
