# Docs for unreleased changes go in a "Next (unreleased)" GitBook variant

**Status:** Accepted (2026-11-20; see [#238](https://github.com/saveenr/VisioAutomation/issues/238) for the discussion). Adopted after a working trial; the variant is public (see Consequences).

## Context

The .NET docs (`VisioAutomation_GitBook_Docs`) publish when a PR merges to `main`. That left no way to write or review docs for a change that is merged in the library but not yet released: either the docs wait until release day, or they go live before the code they describe. Several changes under `[Unreleased]` (the form model and Developer group removals, the renamed `client.Model` methods, the XML page border attributes, `-DirectedGraphStyling`) need docs before they ship.

[#238](https://github.com/saveenr/VisioAutomation/issues/238) listed the options. The two real ones:

- **A separate GitBook variant** backed by its own space, synced to a long-lived `next` branch.
- **Unreleased pages on `main`** with an "Unreleased" note on each.

## Decision

Use a variant. The VisioAutomation site (owned by the `saveenr` GitBook account) has a second variant, **Next (unreleased)**, with its own space. That space has Git Sync to `saveenr/VisioAutomation_GitBook_Docs`, branch `next`, set up with GitHub to GitBook as the initial direction (safe because the space was empty).

Workflow:

1. Docs for an unreleased change are a PR against `next`, not `main`.
2. On release, `next` is merged into `main`, which publishes the released docs.
3. `next` is kept level with `main` between releases (merge `main` into `next` after a docs-only fix on `main`).

Pages that describe unreleased behavior still say so ("Unreleased: added after NuGet 4.0.0"), so a reader who finds the variant is not misled.

## Why

- **The released docs stay correct.** The trial showed the released site unchanged after a PR merged to `next`.
- **It matches the code's own workflow:** an `[Unreleased]` section that is promoted on release.
- **Visibility is not a problem here.** The variant does not need to be hidden; anyone who finds it is told it is unreleased.
- **Cleaner than per-page flags.** No pages to un-flag at release, and no risk of an "unreleased" note outliving its release.

## Consequences

- **Verified on 2026-11-20:** a PR merged into `next` appeared in the Next variant after sync; the released site still showed the old page. While the variant is in Draft it is not listed in the variant picker.
- **PR previews work (verified 2026-11-20):** a throwaway PR against `next` got a GitBook status check linking to a diff in the Next space (opening it needs a GitBook account, as with the PowerShell docs).
- **Undrafted and public (2026-11-20):** the variant is live at `saveenr.gitbook.io/visioautomation/next-unreleased/` with the unreleased text; the released pages are unchanged and the plan (Basic) still serves it. The variant picker is the breadcrumb link above the page title (it shows the current variant name); clicking it lists "Next (unreleased)" alongside the released variant. An earlier version of this record said there was no picker; that was wrong, because only the header and sidebar had been checked. Site search is scoped to the current variant by default ("All content" is opt-in), so unreleased pages do not show up in normal searches. No separate header link is added (decided 2026-11-20); the breadcrumb picker is enough. Draft state hiding the variant from the public site is observed behavior, not documented by GitBook.
- **The variant slug is `next-unreleased`.** Change it if a plain `next` is wanted.
- **Release step added:** merging `next` into `main` becomes part of the release checklist, after the NuGet and module releases.
- **The PowerShell docs use the same approach (2026-11-20).** `VisioPowerShellDocs` got a `next` branch cut from `visiops_v4_docs` (the branch behind the published variant, labelled "Version 4" when this was written and renamed "Version 5" on 2026-11-20 because the module reached 5.x; `master` is not used for publishing) and a "Next (unreleased)" variant on the VisioPowerShell site, with Git Sync from `next` set up GitHub to GitBook on a new blank space. That site already had Version 3 and Version 4 variants, and the picker lists all three. Verified with a real change: PowerShell docs PR 10 (`-DirectedGraphStyling`, marked unreleased) merged into `next` appeared in the Next variant and not in Version 4; GitBook gave the PR a public preview URL. The public pages lagged the completed sync by a minute or two. Release step there: merge `next` into the new version-pinned branch (or `visiops_v4_docs` for a same-major release), then start the next `next` from it. A new major version also needs a new variant, as before.
- **Account dependency:** the site and its Git Sync belong to `saveenr`, not `SevenPens`; the identity moves (#146, #147) must carry this over.

## Cross-references

- [#238](https://github.com/saveenr/VisioAutomation/issues/238) (this decision), [PR 15 on the docs repo](https://github.com/saveenr/VisioAutomation_GitBook_Docs/pull/15) (the trial change).
- [`docs/futures/docs.md`](../futures/docs.md): the "Models docs follow-ups" backlog entry, whose pages now have somewhere to go.

## Reconsider when

- Readers do not find the breadcrumb picker, or GitBook changes search scoping.
- GitBook plan or pricing changes remove variants from our plan.
- Keeping `next` level with `main` becomes a recurring chore.
