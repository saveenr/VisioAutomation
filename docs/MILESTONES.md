# MILESTONES.md &mdash; planning by semester

Forward-looking work plan organized by **semester** (calendar year + quarter, e.g. `CY26Q2`). Sits between [`ROADMAP.md`](ROADMAP.md) (high-level phase plan) and [`FUTURES.md`](FUTURES.md) (backlog index, no scheduling).

The semester is the primary planning unit. Each semester maps to a [GitHub milestone](https://github.com/saveenr/VisioAutomation/milestones) of the same name. Themed groupings ("Milestones A, B, C, ...") stay as stable identifiers for *what kind of work* an item is, while the *when* is the semester.

> Done milestones (Phase 1, Phase 2, etc.) are summarized in [`COMPLETED.md`](COMPLETED.md). What follows is forward-looking only.

---

## Guiding principle: improve before audience-reducing changes

Before any change that would reduce the project's reach &mdash; TFM bumps, modern-.NET moves, Visio version baseline shifts &mdash; we prioritize work that makes the *current* audience better served: documentation, convenience features, code quality, testing, maintainability, and addressing user feedback. The reasoning: if anyone later wants to pick up maintenance of an older-Visio or older-.NET fork, they should start from a finished, well-documented, well-tested baseline rather than inheriting a half-renovated codebase.

This principle drives the semester sequencing below. **2026 (CY26Q2 through CY26Q4)** is dedicated to audience-preserving improvements. **2027 onward (CY27Q1+)** is when audience-reducing modernization begins, after the Windows 10 LTSB 2016 mainstream-support sunset (2026-10-13) and after the 2026 priority work is mature.

---

## Status as of 2026-11-20

> **Draft refresh, pending owner review.** The schedule below moves several items between semesters; those moves are proposals until the GitHub milestones are re-tagged to match.

- **CY26Q2 is complete.** All 14 issues in the milestone are closed.
- **Releases since this doc was last updated (2026-05-08):** NuGet 3.2.0 and Visio PowerShell 4.8.0 on 2026-10-01; NuGet 4.0.0 and Visio PowerShell 5.0.0 on 2026-11-01. Both 2026-11-01 releases raise the minimum runtime to .NET Framework 4.7.2. The version-number policy question was re-evaluated in May and resolved as "stay divergent" ([`futures/releases.md`](futures/releases.md)).
- **Milestone B (modernization unlock) landed early.** [#171](https://github.com/saveenr/VisioAutomation/issues/171) (VS 2026) and [#174](https://github.com/saveenr/VisioAutomation/issues/174) (net472) closed on 2026-11-01, after the Windows 10 LTSB 2016 sunset on 2026-10-13. They were scheduled for CY27Q1.
- **CY26Q3 is past due.** Seven of its issues are still open: [#133](https://github.com/saveenr/VisioAutomation/issues/133), [#155](https://github.com/saveenr/VisioAutomation/issues/155), [#157](https://github.com/saveenr/VisioAutomation/issues/157), [#162](https://github.com/saveenr/VisioAutomation/issues/162), [#163](https://github.com/saveenr/VisioAutomation/issues/163), [#164](https://github.com/saveenr/VisioAutomation/issues/164), [#173](https://github.com/saveenr/VisioAutomation/issues/173). The proposal below re-homes each.
- **A new class of work appeared that this doc never planned for:** about 20 Models and facade defects (#197 through #232) were fixed and closed on 2026-10-01, and the design questions they exposed remain open as [#218](https://github.com/saveenr/VisioAutomation/issues/218), [#220](https://github.com/saveenr/VisioAutomation/issues/220), [#221](https://github.com/saveenr/VisioAutomation/issues/221), [#223](https://github.com/saveenr/VisioAutomation/issues/223), [#198](https://github.com/saveenr/VisioAutomation/issues/198), [#199](https://github.com/saveenr/VisioAutomation/issues/199) and [#236](https://github.com/saveenr/VisioAutomation/issues/236). They are now Milestone I below.
- **Open issues:** 28, of which 9 have no milestone.

---

## Semester schedule

### CY26Q2 &mdash; April&ndash;June 2026 *(complete)*
[GitHub milestone](https://github.com/saveenr/VisioAutomation/milestone/2) &mdash; due 2026-06-30 &mdash; 14 of 14 closed.

**Theme:** Kick off audience-preserving improvements. Delivered: the PSVA and VisioBot3000 scoping reviews ([#149](https://github.com/saveenr/VisioAutomation/issues/149), [#150](https://github.com/saveenr/VisioAutomation/issues/150)), the portfolio audit Phase H1 ([#152](https://github.com/saveenr/VisioAutomation/issues/152)), the tests-need-Visio decision ([#153](https://github.com/saveenr/VisioAutomation/issues/153)), the version compatibility tables ([#161](https://github.com/saveenr/VisioAutomation/issues/161)), the Models docs ([#132](https://github.com/saveenr/VisioAutomation/issues/132)), the coverage audit ([#154](https://github.com/saveenr/VisioAutomation/issues/154)), and the triage pass ([#151](https://github.com/saveenr/VisioAutomation/issues/151)). [#163](https://github.com/saveenr/VisioAutomation/issues/163) was listed here and has since moved to CY26Q3 and then CY27Q1.

### CY26Q3 &mdash; July&ndash;September 2026 *(past due; 9 closed, 7 open)*
[GitHub milestone](https://github.com/saveenr/VisioAutomation/milestone/3) &mdash; due 2026-09-30

**Theme:** Execute on the May 2026 cmdlet-ergonomics shortlist; close the docs and test decisions that block later work. Closed: [#131](https://github.com/saveenr/VisioAutomation/issues/131), [#156](https://github.com/saveenr/VisioAutomation/issues/156), [#176](https://github.com/saveenr/VisioAutomation/issues/176), [#177](https://github.com/saveenr/VisioAutomation/issues/177), [#178](https://github.com/saveenr/VisioAutomation/issues/178), [#181](https://github.com/saveenr/VisioAutomation/issues/181), [#182](https://github.com/saveenr/VisioAutomation/issues/182), [#184](https://github.com/saveenr/VisioAutomation/issues/184), [#191](https://github.com/saveenr/VisioAutomation/issues/191).

| Open item | Themed milestone | Issue | Proposed new home |
|---|---|---|---|
| Add troubleshooting page to .NET gitbook | E | [#133](https://github.com/saveenr/VisioAutomation/issues/133) | CY26Q4 |
| Decide: long-term docs location | E | [#157](https://github.com/saveenr/VisioAutomation/issues/157) | CY26Q4 (gates the docs restructure) |
| Discussion: doc-sample-as-test linkage mechanism | E | [#155](https://github.com/saveenr/VisioAutomation/issues/155) | CY27Q1 (follows #157) |
| Cmdlet-binding test coverage | F | [#173](https://github.com/saveenr/VisioAutomation/issues/173) | CY26Q4 (continuing; about 40 cmdlets still untested) |
| Investigate switching `VisioCmdlet` to `PSCmdlet` | D | [#164](https://github.com/saveenr/VisioAutomation/issues/164) | CY27Q1 |
| Pipeline parameter set on `Connect-VisioShape` | D | [#163](https://github.com/saveenr/VisioAutomation/issues/163) | CY27Q1 (blocked by #164) |
| Phase A scoping: PSVA candidate cmdlets | D | [#162](https://github.com/saveenr/VisioAutomation/issues/162) | CY27Q1 |

### CY26Q4 &mdash; October&ndash;December 2026 *(current)*
[GitHub milestone](https://github.com/saveenr/VisioAutomation/milestone/1) &mdash; due 2026-12-31 &mdash; 1 closed ([#190](https://github.com/saveenr/VisioAutomation/issues/190), 2026-11-20), 5 open.

**Theme:** Identity transition push, plus settling the public Models and facade API before more docs are written against it. The identity items are owner tasks rather than code.

| Item | Themed milestone | Issue |
|---|---|---|
| Migrate GitHub repo to SevenPens-owned account | A | [#146](https://github.com/saveenr/VisioAutomation/issues/146) |
| Migrate gitbook spaces to SevenPens-owned hosting | A | [#147](https://github.com/saveenr/VisioAutomation/issues/147) |
| Phase 5b: in-repo URL rewrite (single commit after #146 + #147 land) | A | (no issue; one-shot commit) |
| Retire unused `VisioAutomation` PSGallery co-owner account | A | [#148](https://github.com/saveenr/VisioAutomation/issues/148) (needs a milestone) |
| Open NuGet support case to declassify the `saveenr` account | A | [#179](https://github.com/saveenr/VisioAutomation/issues/179) |
| New logo (refresh, XAR to SVG); [#80](https://github.com/saveenr/VisioAutomation/issues/80) ruled that branding belongs here | A | [#196](https://github.com/saveenr/VisioAutomation/issues/196) (needs a milestone) |
| Decide Box layout, form page model and `ContainerLayout` together | I | [#218](https://github.com/saveenr/VisioAutomation/issues/218), [#220](https://github.com/saveenr/VisioAutomation/issues/220), [#221](https://github.com/saveenr/VisioAutomation/issues/221) (need milestones) |
| Revise user-facing gitbook docs for accuracy | E | [#172](https://github.com/saveenr/VisioAutomation/issues/172) |
| Investigate an unreleased (`next`) docs variant | E | [#238](https://github.com/saveenr/VisioAutomation/issues/238) |
| Troubleshooting page and docs-location decision | E | [#133](https://github.com/saveenr/VisioAutomation/issues/133), [#157](https://github.com/saveenr/VisioAutomation/issues/157) (moved from CY26Q3) |
| Cmdlet-binding test coverage, next slice | F | [#173](https://github.com/saveenr/VisioAutomation/issues/173) (moved from CY26Q3) |

### CY27Q1 &mdash; January&ndash;March 2027
[GitHub milestone](https://github.com/saveenr/VisioAutomation/milestone/4) &mdash; due 2027-03-31 &mdash; 2 closed early, 4 open.

**Theme:** The modernization unlock is done ([#171](https://github.com/saveenr/VisioAutomation/issues/171) and [#174](https://github.com/saveenr/VisioAutomation/issues/174) shipped 2026-11-01). What remains is spike and audit work, the cmdlet-ergonomics follow-through that slipped from CY26Q3, and the Models design discussions that do not need a decision this year.

| Item | Themed milestone | Issue |
|---|---|---|
| Spike: does the existing Visio PIA work for COM interop on .NET 11? (time-boxed, throwaway branch; no NetOffice) | C | [#249](https://github.com/saveenr/VisioAutomation/issues/249) |
| Audit: identify Visio-2010-only paths in the codebase | G | [#159](https://github.com/saveenr/VisioAutomation/issues/159) |
| ShapeSheet query ergonomics spike (LINQ and other shapes) | F | [#170](https://github.com/saveenr/VisioAutomation/issues/170) |
| `VisioScripting` dead-method removal (re-verify zero callers first) | B | [#183](https://github.com/saveenr/VisioAutomation/issues/183) |
| `PSCmdlet` investigation, then `Connect-VisioShape` pipeline set | D | [#164](https://github.com/saveenr/VisioAutomation/issues/164), [#163](https://github.com/saveenr/VisioAutomation/issues/163) (moved from CY26Q3) |
| Phase A scoping: PSVA candidate cmdlets | D | [#162](https://github.com/saveenr/VisioAutomation/issues/162) (moved from CY26Q3) |
| Doc-sample-as-test linkage | E | [#155](https://github.com/saveenr/VisioAutomation/issues/155) (moved from CY26Q3) |
| `client.Model` draw method consistency; directed graph XML gaps; single directed graph on the current page | I | [#223](https://github.com/saveenr/VisioAutomation/issues/223), [#236](https://github.com/saveenr/VisioAutomation/issues/236), [#198](https://github.com/saveenr/VisioAutomation/issues/198), [#199](https://github.com/saveenr/VisioAutomation/issues/199) (need milestones) |

### CY27Q2 &mdash; April&ndash;June 2027
[GitHub milestone](https://github.com/saveenr/VisioAutomation/milestone/5) &mdash; due 2027-06-30 &mdash; 3 open.

**Theme:** Visio baseline and portfolio decisions, informed by the CY27Q1 spike and audit.

| Item | Themed milestone | Issue |
|---|---|---|
| Move the libraries to .NET 11 and add a PowerShell 7 module (multi-target alongside net472 if the CY27Q1 spike allows); [#175](https://github.com/saveenr/VisioAutomation/issues/175) was retitled for this on 2026-11-20 | C | [#175](https://github.com/saveenr/VisioAutomation/issues/175) |
| Decide: Visio 2013 baseline branding | G | [#160](https://github.com/saveenr/VisioAutomation/issues/160) |
| Visio 2013 baseline migration implementation (gated on #160) | G | (no issue yet) |
| Decide: consolidate Visio-PIAs source into this repo | H | [#180](https://github.com/saveenr/VisioAutomation/issues/180) |

### Beyond CY27Q2

Items pending but not yet semester-assigned:
- **Long-term investigation: replace the Visio PIA with NetOffice / NetOfficeFw** ([#158](https://github.com/saveenr/VisioAutomation/issues/158), retitled and milestone removed 2026-11-20). Not part of the plan. It stays on the back burner until something forces it, for example the existing Visio PIA not working on modern .NET.
- Documentation work past the CY26Q4 burst &mdash; continuous into 2027 as items come up.
- Visio repo portfolio audit Phase H2 (per-repo retire / merge / maintain decisions). Opportunistic.

---

## Themed milestones (the *what kind of work* axis)

The semester schedule above is the *when*. The themed milestones below are the *what kind of work*. Each item in the schedule is tagged with a milestone letter so you can read by theme. Each milestone may span multiple semesters.

### Milestone A &mdash; Identity transition complete
**Theme:** Finish the dev-team identity transition from `Saveen` / `saveenr` to `SevenPens`.

**User impact:** Old hosting URLs keep resolving via redirects; new canonical URLs live under `SevenPens`-owned hosting. No broken bookmarks, no consumer disruption.

**Semesters:** CY26Q4 (target window).

**Items:** see CY26Q4 schedule above. Detail in [`futures/identity.md`](futures/identity.md).

### Milestone B &mdash; Modernization unlock
**Theme:** Bump shipping-lib TFMs from .NET Framework 4.5.2 to 4.7.2, move dev to VS 2026, clear longstanding deprecation / version-hygiene cruft.

**User impact:** *Audience-reducing.* Minimum supported .NET Framework rises from 4.5.2 to 4.7.2 &mdash; affects only consumers on pre-1803 Windows, which leave mainstream support October 2026. Modern IDE tooling, faster builds, cleaner manifest metadata.

**Semesters:** CY27Q1 (target window). Gated on the Windows 10 LTSB 2016 sunset 2026-10-13.

**Status (2026-11-20): done, ahead of schedule.** The TFM bump and the move to VS 2026 shipped on 2026-11-01 in NuGet 4.0.0 and Visio PowerShell 5.0.0 ([#171](https://github.com/saveenr/VisioAutomation/issues/171), [#174](https://github.com/saveenr/VisioAutomation/issues/174)). The manifest cleanups and Release builds landed earlier ([#177](https://github.com/saveenr/VisioAutomation/issues/177), [#178](https://github.com/saveenr/VisioAutomation/issues/178), [#181](https://github.com/saveenr/VisioAutomation/issues/181)). Only [#183](https://github.com/saveenr/VisioAutomation/issues/183), the `VisioScripting` dead-method removal, remains under this theme.

**Items:** see CY27Q1 schedule above. Detail in [`futures/build-and-code.md`](futures/build-and-code.md) and [`futures/releases.md`](futures/releases.md).

### Milestone C &mdash; Modern .NET (multi-target)
**Theme:** Move to .NET 11 and add a PowerShell 7 module, after the feature, test and docs work on the current line. Direction set 2026-11-20, prompted by Microsoft's announcement that PowerShell 7 will ship with the next release of Windows. Replacing the Visio PIA with NetOffice / NetOfficeFw is **not** part of this milestone; it is a long-term investigation (see "Beyond CY27Q2").

**User impact:** Modern C# features and better modern-.NET performance become available. Existing `net472` consumers are unaffected if the libraries multi-target. Windows PowerShell 5.1 users of VisioPS keep the `net472` build; PowerShell 7+ users get the .NET 11 build. If multi-targeting turns out not to work, the fallback is a separate new line with a maintenance branch for the current one. Windows PowerShell 5.1 stays in use for a long time, so the current line keeps getting bug fixes.

**Semesters:** CY27Q1 (a time-boxed spike on whether the existing Visio PIA works for COM interop on .NET 11) &rarr; CY27Q2 (implementation begins). Sequenced after Milestone B.

**Items:** see CY27Q1 and CY27Q2 schedules above. Detail in [`futures/build-and-code.md`](futures/build-and-code.md).

### Milestone D &mdash; Cmdlet ergonomics
**Theme:** Adopt the good ideas from sibling community PowerShell-for-Visio projects (VisioBot3000, PSVA) plus the `EncodeValues` simplification.

**User impact:** Friendlier authoring for VisioPS users &mdash; nickname registry, block-style nesting for containers, bulk shape operations, custom-property authoring without manual `EncodeValues()`. Limited to the items the May 2026 scoping review picks; rest stays in backlog.

**Semesters:** CY26Q2 (review) &rarr; CY26Q3 (implementation through bandwidth).

**Items:** see CY26Q2 and CY26Q3 schedules above. Detail in [`futures/build-and-code.md`](futures/build-and-code.md).

### Milestone E &mdash; Documentation completeness
**Theme:** Close documentation gaps on user-facing gitbooks and in-repo dev guides. Includes the long-running policy decisions about *where* docs live and *what* is public.

**User impact:** Easier onboarding for new contributors and consumers; fewer "had to read the source" moments. Reference docs cover the surface that users actually program against.

**Semesters:** CY26Q3 (concentrated burst &mdash; the existing docs issues + the location/public-API decisions). Continuous through CY26Q4 and into 2027 as items come up.

**Items:** see CY26Q3 schedule above. Detail in [`futures/docs.md`](futures/docs.md).

**Plus longer-running items not yet semester-assigned** (will be picked up as bandwidth allows):
- Restructure user-docs repos &mdash; gated on the docs-location decision ([#157](https://github.com/saveenr/VisioAutomation/issues/157)).
- Keep CHANGELOGs current (process item, not a one-shot).
- **Visio version &harr; PIA mapping reference page** &mdash; new doc page explaining the version-number table (Visio 2010 = 14, 2013 = 15, ...) and where to get each PIA. Couples to the [`Visio-PIAs`](https://github.com/saveenr/Visio-PIAs) sibling repo.

### Milestone F &mdash; Test infrastructure
**Theme:** Modernize the test suite and address the live-Visio dependency that currently keeps tests off CI.

**User impact:** More reliable releases (catch regressions in CI rather than at publish time); contributor onboarding without a Visio install for the test buckets that don't actually need one.

**Semesters:** CY26Q2 (design-decision write-up) &rarr; CY26Q3 (coverage audit) &rarr; CY26Q4+ (implementation as bandwidth allows).

**Items:** see CY26Q2 and CY26Q3 schedules above. Detail in [`futures/build-and-code.md`](futures/build-and-code.md) and [`futures/tests.md`](futures/tests.md). The "tests need a live Visio" rule that gates CI today is now formalized in [`decisions/tests-need-visio.md`](decisions/tests-need-visio.md). The no-Visio test split has *de facto* started: [`ManifestTests`](../VisioAutomation_2010/VTest.PowerShell/ManifestTests.cs) and [`XmlErrorLogTests`](../VisioAutomation_2010/VTest/Core/Application/XmlErrorLogTests.cs) are existing examples of the no-Visio bucket; the planned tagging work to turn the implicit split explicit is part of this milestone.

### Milestone G &mdash; Visio 2013 baseline migration
**Theme:** Move the codebase's baseline Visio version from 2010 to 2013.

**User impact:** *Audience-reducing.* Consumers still on Visio 2010 (released 2010, mainstream support ended 2015, extended support ended 2020) would no longer see new releases. Visio 2013+ users get access to capabilities the 2013-era Visio API added.

**Semesters:** CY27Q1 (audit) &rarr; CY27Q2 (branding decision &rarr; implementation).

**Items:** see CY27Q1 and CY27Q2 schedules above.

### Milestone H &mdash; Visio repo portfolio audit
**Theme:** Take stock of the broader Visio-related repo portfolio. Make each repo's status visible; decide what to retire / merge / maintain.

**User impact:** Visitors to any of the 9 sibling repos can tell at a glance whether it's active. Reduces "is this still maintained?" friction. Cleanup of dormant repos shrinks reader surface area.

**Semesters:** CY26Q2 (Phase H1: index + per-repo READMEs). CY26Q3+ for Phase H2 (per-repo retire/merge/maintain decisions, opportunistic).

**Status (2026-11-20):** Phase H1 is done ([#152](https://github.com/saveenr/VisioAutomation/issues/152) closed 2026-05-08). Phase H2 is unscheduled; the one open piece is [#180](https://github.com/saveenr/VisioAutomation/issues/180) (consolidate Visio-PIAs), in CY27Q2. `visio-templates` was deleted, so the list below has one fewer live repo.

**Items:** see CY26Q2 schedule. Phase H2 not yet scheduled to a specific semester. Repos in scope:

- [`saveenr/VisioAutomation`](https://github.com/saveenr/VisioAutomation) (this repo) &mdash; primary, active.
- [`saveenr/VisioAutomation2007`](https://github.com/saveenr/VisioAutomation2007)
- [`saveenr/VisioAutomation.VDX`](https://github.com/saveenr/VisioAutomation.VDX)
- [`saveenr/Visio-PIAs`](https://github.com/saveenr/Visio-PIAs)
- [`saveenr/visio-templates`](https://github.com/saveenr/visio-templates)
- [`saveenr/visio-reference`](https://github.com/saveenr/visio-reference)
- [`saveenr/Visio-Power-Tools`](https://github.com/saveenr/Visio-Power-Tools)
- [`saveenr/Visio-Export-Pages-To-Docs`](https://github.com/saveenr/Visio-Export-Pages-To-Docs)
- [`saveenr/Visio-Font-Compare`](https://github.com/saveenr/Visio-Font-Compare)
- [`saveenr/Visio-Code-Samples`](https://github.com/saveenr/Visio-Code-Samples)

### Milestone I &mdash; Models and facade API cleanup *(new, 2026-11-20)*
**Theme:** Settle the public surface of `VisioAutomation.Models` and the `client.Model` draw methods. About 20 defects in this area were fixed on 2026-10-01 (#197 through #232); the questions they exposed are design decisions, not bugs, and several docs pages describe behavior that a decision may change.

**User impact:** Consumers of the diagram models get a smaller, consistent surface. Some outcomes may remove or reshape public types ([#218](https://github.com/saveenr/VisioAutomation/issues/218) Box layout, [#220](https://github.com/saveenr/VisioAutomation/issues/220) form page model, [#221](https://github.com/saveenr/VisioAutomation/issues/221) `ContainerLayout`), which is a compatibility decision under the [`VisioScripting` public-API contract](decisions/visioscripting-public-api.md) and the changelog rules.

**Semesters:** CY26Q4 for the three type decisions, which should be taken together (the docs side of each is listed in [`futures/docs.md`](futures/docs.md)); CY27Q1 for [#223](https://github.com/saveenr/VisioAutomation/issues/223), [#236](https://github.com/saveenr/VisioAutomation/issues/236), [#198](https://github.com/saveenr/VisioAutomation/issues/198) and [#199](https://github.com/saveenr/VisioAutomation/issues/199).

**Items:** see CY26Q4 and CY27Q1 schedules above. None of these seven issues has a GitHub milestone yet.

---

## Triage backlog (handled in [#151](https://github.com/saveenr/VisioAutomation/issues/151), CY26Q2)

> **All resolved as of 2026-11-20.** [#82](https://github.com/saveenr/VisioAutomation/issues/82) was fixed and closed on 2026-05-07, and [#117](https://github.com/saveenr/VisioAutomation/issues/117) was closed on 2026-10-01, which also closed [#151](https://github.com/saveenr/VisioAutomation/issues/151). The notes below are the original 2026-05-07 outcomes, kept for the record.

Older issues triaged on 2026-05-07. Per-issue rationale in [#151's decision comment](https://github.com/saveenr/VisioAutomation/issues/151#issuecomment-4401101627). Summary of outcomes:

- [#80](https://github.com/saveenr/VisioAutomation/issues/80) &mdash; New logo. **Closed as not-planned** (any branding work belongs in Axis 5, CY26Q4).
- [#82](https://github.com/saveenr/VisioAutomation/issues/82) &mdash; Size + Cells on directed graph node. **Real bug, fix in flight on a branch this session.** Milestoned to CY26Q2; closes when the fix lands on master.
- [#102](https://github.com/saveenr/VisioAutomation/issues/102) &mdash; Problem creating shape from template. **Closed as completed** (both questions answered in-thread on 2026-05-06).
- [#105](https://github.com/saveenr/VisioAutomation/issues/105) &mdash; directed-graph layout/direction umbrella. **Closed as completed** (sub-items shipped in NuGet 3.0.0; docs page live).
- [#117](https://github.com/saveenr/VisioAutomation/issues/117) &mdash; Custom-properties on directed graph. **Open, revisit ~2026-05-27.** Fix shipped in [Visio PowerShell 4.7.0](https://github.com/saveenr/VisioAutomation/releases/tag/VisioPS_4.7.0) on 2026-05-06; awaiting reporter confirmation.

---

## Updating this doc

- When a milestone item lands, move it from this doc to [`COMPLETED.md`](COMPLETED.md) (or strike it through here if the semester is mid-flight).
- When a proposed semester slips, move the item to a later semester in this doc and re-tag the GitHub issue's milestone.
- When a new item arrives, decide its semester first (when does it land?), then its themed milestone (what kind of work is it?). Add to the semester schedule and tag the issue accordingly.
- Add new semesters as the long-term picture clarifies. CY27Q3 and beyond aren't yet structured here; they will be when the current 2027 work matures enough to forecast.
- The targets here are not commitments to consumers &mdash; release notes and CHANGELOGs are the authoritative "what's shipped when" record. This doc is internal planning.
- The guiding principle (improve before audience-reducing changes) governs sequencing. If a future change bumps a 2027+ item earlier, the principle should be revisited at the same time, not silently overruled.
