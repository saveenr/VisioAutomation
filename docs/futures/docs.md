# Futures — Documentation

Backlog of documentation items, both in-repo developer docs and the user-facing gitbook docs. Index of all backlog files: [`../FUTURES.md`](../FUTURES.md).

---

### Decide where docs live long-term
- **What:** User docs are in a separate repo on gitbook ([`VisioAutomation_GitBook_Docs`](https://github.com/saveenr/VisioAutomation_GitBook_Docs)); developer docs are now in `docs/` here.
- **Why:** Two-repo doc setups drift. Either keep them split with a clear policy (which doc lives where) or consolidate. No urgent action needed — just call out the policy in `OVERVIEW.md` once decided.
- **Effort:** S (policy) — or M (consolidation).
- **Cross-refs:** *Restructure the user-docs repos* below covers the related-but-distinct question of how the two user-doc gitbooks are themselves arranged.

### Restructure the user-docs repos
- **What:** User-facing docs currently live in **two** separate gitbook repos: [`VisioAutomation_GitBook_Docs`](https://github.com/saveenr/VisioAutomation_GitBook_Docs) (.NET, on `main`) and [`VisioPowerShellDocs`](https://github.com/saveenr/VisioPowerShellDocs) (PowerShell, on a version-pinned `visiops_v4_docs` branch). Three repos total counting the code repo. Question: is this the right shape?
- **Why this came up:** The 2026-05 doc-audit work (compile-checking C# snippets via VPlayground, parser-checking PowerShell blocks across both gitbooks, fanning out atomic doc-fix commits to multiple remotes) made the cross-repo workflow visibly expensive. `CLAUDE.md` already flags "Two-repo doc setups drift" as a known concern. The recently-filed [#131](https://github.com/saveenr/VisioAutomation/issues/131) / [#132](https://github.com/saveenr/VisioAutomation/issues/132) (large doc-coverage issues) and [#133](https://github.com/saveenr/VisioAutomation/issues/133) (troubleshooting page) will all be more painful in a multi-repo setup.

#### Options surveyed (2026-05-05 discussion)

1. **Status quo — 3 repos.** Code in `VisioAutomation`; .NET docs in `VisioAutomation_GitBook_Docs`; PS docs in `VisioPowerShellDocs`. Most cross-repo friction. Highest drift risk (3 places to keep in sync).
2. **Orphan branches in the code repo.** Add `docs-dotnet` / `docs-powershell-v4` orphan branches to `VisioAutomation`. One repo total. Removes a remote, but still loses code+doc atomic commits (different branches), and code-repo workflows (`git log`, GitHub UI) gain noise from doc churn.
3. **Subdirectory on code-repo master.** A `docs-site/dotnet/` and `docs-site/powershell/` tree on `master`. One repo, one branch, one log. **Enables atomic code+doc commits** — the structural fix for drift. CI could compile-check doc snippets against current source. Heaviest change to existing workflows; biggest payoff.
4. **Single separate docs repo, orphan branches per doc set.** Collapse the two doc repos into one (e.g. `VisioAutomation_Docs` with `dotnet` and `powershell-v4` orphan branches). Cuts 3 repos to 2. Cheapest migration (~2 hours). Doesn't fix atomic commits or audit friction — the structural doc-drift mechanic is unchanged, just spread across fewer remotes.

#### Decision factors

- **What kind of doc work do we actually do?** If most doc work is "doc-only sessions" like the 2026-05 audit, options 1 and 4 are tolerable. If most doc work is "code change + doc update in same session", only option 3 actually helps.
- **How much do we care about CI compile-checking doc snippets?** Today the audits do it manually via VPlayground. Automating it requires the source to be in the same checkout as the docs. Option 3 enables it cheaply; the others require cross-repo CI.
- **How much do we value clean separation of code and docs in `git log` / GitHub UI?** If high, options 1 / 2 / 4 win. If low, option 3 is the simplest model.
- **Is GitBook config flexible enough for any of these?** Yes — GitBook can read any branch + any subpath of any repo. None of the options is mechanically blocked by the publishing platform.

#### Migration cost

- Option 1 → option 4: **~2 hours.** Rename one doc repo, push the other's content as a new orphan branch, repoint one GitBook space, archive the now-redundant repo.
- Option 1 → option 3: **half a day to a full day.** Move both doc trees into `docs-site/` subdirs on master, repoint both GitBook spaces, archive both old repos, update `CLAUDE.md` and `reference_doc_repos.md` memory entries.
- Option 1 → option 2: similar to option 3 minus the subdir-vs-orphan-branch difference.

#### Status

- **Held for further discussion.** Not blocking; the status quo works, just expensive. Tackle when the next big doc-audit or cross-cutting code+doc change makes the cost concrete again.
- **Forcing function:** there isn't one; doc structure can change at any time. But aligning with a NuGet/PSGallery release would be a natural moment, since that's already a coordinated cross-product event.

#### Cross-refs

- *Decide where docs live long-term* (above) — the dev-docs-vs-user-docs policy question. This entry is about the user-docs side specifically.
- [#131](https://github.com/saveenr/VisioAutomation/issues/131), [#133](https://github.com/saveenr/VisioAutomation/issues/133), [#172](https://github.com/saveenr/VisioAutomation/issues/172) — large doc-coverage work that will benefit from whichever structure is chosen. ([#132](https://github.com/saveenr/VisioAutomation/issues/132) closed; the original `VisioAutomation.Models` Tier 3 audit shipped 2026-05-07.)

#### Effort

- S for option 4 (~2 hours).
- M for option 3 or option 2 (half a day to a full day).
- N/A (no work) for option 1.

### Decide whether to document `VisioScripting` as a public API
- **What:** `VisioScripting` is the .NET layer between the PowerShell cmdlets and the underlying `VisioAutomation` library. Its `Client` object groups commands by topic (`Document`, `Page`, `Selection`, `View`, `Text`, `Shape`, `ShapeSheet`, `Application`, `Master`, `Container`, `Connection`, `Hyperlink`, `Lock`, `CustomProperty`, `UserDefinedCell`, `Output`, `Undo`, `Window`, `Layer`, `Color`, etc.) — most cmdlets are thin wrappers over a `Client.<Group>.<Method>(...)` call.
- **External feedback (2026-05-05):** A doc-review pass on the gitbook ([proposed-issues.md issue #1](https://github.com/saveenr/VisioAutomation_GitBook_Docs/blob/main/proposed-issues.md), since converted into a GitHub issue) flagged the `VisioScripting` gap as the single highest-priority documentation hole. The argument: the source [`readme.md`](../../readme.md) leads with a `VisioScripting.Client` snippet as its quick-start, so a new C# reader's *first impression* is an undocumented type. That moves the open question below ("part of the project's promised surface, or internal?") from theoretical to forcing-function — answer it before deciding whether to fill the gap.
- **Currently documented:** only as power-user escape hatches. The PS-side `cmdlets/other-cmdlets.md` lists `Get-VisioClient` (which returns a `VisioScripting.Client`); `technical-notes/getting-the-current-scriptingsession.md` and `technical-notes/use-visioautomation.md` give brief pointers to the .NET-side bridge. There is no per-method or per-group reference for `VisioScripting` itself.
- **Why this is a real question, not just a coverage gap:**
  - **Audience.** `VisioScripting` is a *higher-level* alternative to the raw `VisioAutomation` library — you'd reach for it from .NET when you want commands like "duplicate this page" or "select all shapes" without composing them yourself from `Page.Pages.Add` + `ShapeSheet.Writers.SrcWriter` + ... . That's a real audience, separate from PowerShell users.
  - **Stability.** Right now `VisioScripting` is treated as an internal implementation detail of the cmdlets — APIs may shift to suit cmdlet needs. Documenting it elevates it to a public surface, which changes the cost of API churn.
  - **Surface size.** Roughly one Helper / Commands class per topic, each with 5–20 methods. Order-of-magnitude similar to the .NET-side Tier 1+2+4 work that was just done (~15 pages).
- **Decisions to make first:**
  - **Is `VisioScripting` part of the project's promised surface, or an internal that shouldn't be relied on?** Affects whether documentation should exist at all and whether the cmdlets should keep wrapping it.
  - **Same gitbook or separate?** Could be a third gitbook, or a section under [VisioAutomation_GitBook_Docs](https://github.com/saveenr/VisioAutomation_GitBook_Docs).
- **Cross-refs:** *Decide where docs live long-term* (related policy question). The recently-completed Tiers 1+2+4 .NET-side coverage work is the validated pattern this would follow.
- **Effort:** S to decide. M–L to write if the answer is "yes, document it" (similar in size to Tiers 1+2+4 of the .NET-side coverage).

### Keep CHANGELOGs current as Phase 1 work lands
- **What:** Two changelogs were added in [Keep a Changelog](https://keepachangelog.com/en/1.1.0/) format: [`NuGet/CHANGELOG.md`](../../NuGet/CHANGELOG.md) for the `VisioAutomation2010` NuGet, and [`VisioAutomation_2010/VisioPowerShell/CHANGELOG.md`](../../VisioAutomation_2010/VisioPowerShell/CHANGELOG.md) for the `Visio` PowerShell module. Each has an `[Unreleased]` section that should accumulate consumer-visible changes until the Phase 2 release cuts a real version.
- **Why:** The whole point of cutting a final release in Phase 2 is to give consumers a clean, well-documented checkpoint. If Unreleased sections drift behind reality during Phase 1, the release notes will be wrong.
- **How to apply:** When a Phase 1 commit changes anything a consumer of the NuGet or PS module would notice (public API, parameter behavior, supported runtime, dependencies), add an entry to the corresponding CHANGELOG's `[Unreleased]` in the same commit. Pure internal/build/docs changes don't need entries.
- **Effort:** ~zero per change, if done in the same commit.

### Add a troubleshooting page to the .NET gitbook
- **What:** Neither gitbook has a Troubleshooting / FAQ page. Surfaced by the 2026-05-05 doc-review pass ([proposed-issues.md issue #8](https://github.com/saveenr/VisioAutomation_GitBook_Docs/blob/main/proposed-issues.md)) which sketched the candidate failure modes: COM-registration failures when Visio isn't installed; PIA-version vs. `VisioAutomation2010`-version mismatches; stencil-filename differences across Visio versions; 32-bit vs. 64-bit PowerShell host with the `Visio` module; "failed to log in to github.com" errors when publishing.
- **Why deferred (not in Group B):** speculatively-written troubleshooting pages age badly and tend to confuse more than help. Better to wait until we have a real corpus of user-reported failures to ground the page in. The candidate list above is the seed.
- **How to apply:** when filing real bug reports / issues, tag those that are environmental ("works on my machine"-class) for inclusion. Build the page reactively from accumulated cases rather than upfront.
- **Effort:** S–M once there's enough real material to justify it.

### Revise user-facing documentation for accuracy
- **What:** Audit the public gitbook docs ([VisioAutomation](https://saveenr.gitbook.io/visioautomation/) and [Visio PowerShell](https://saveenr.gitbook.io/visiopowershell/), source repo: [VisioAutomation_GitBook_Docs](https://github.com/saveenr/VisioAutomation_GitBook_Docs)) against the current API surface. Update or remove anything that no longer matches the code, and fill in coverage for cmdlets / APIs that have been added since the docs were last touched.
- **Why:** The docs have not been refreshed alongside recent changes; users hitting a stale example as their first impression is the worst kind of regression.
- **Approach (suggested):**
  - Start with the **PowerShell module** since it has the most cmdlet-by-cmdlet documentation surface and is the most user-facing.
  - For each cmdlet, verify it still exists, parameters still match, and the example still runs.
  - Do the C# library docs second.
  - Use the new [`docs/ARCHITECTURE.md`](../ARCHITECTURE.md) and [`docs/GLOSSARY.md`](../GLOSSARY.md) as the source of truth for terminology and structure.
- **Cross-refs:** Related to but distinct from *Decide where docs live long-term* — that item is about the gitbook-vs-in-repo *policy*; this item is about *accuracy of the existing user-facing content*.
- **Effort:** L (the cmdlet inventory alone is substantial).
