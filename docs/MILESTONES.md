# MILESTONES.md &mdash; planning by semester

Forward-looking work plan organized by **semester** (calendar year + quarter, e.g. `CY26Q2`). Sits between [`ROADMAP.md`](ROADMAP.md) (high-level phase plan) and [`FUTURES.md`](FUTURES.md) (backlog index, no scheduling).

The semester is the primary planning unit. Each semester maps to a [GitHub milestone](https://github.com/saveenr/VisioAutomation/milestones) of the same name. Themed groupings ("Milestones A, B, C, ...") stay as stable identifiers for *what kind of work* an item is, while the *when* is the semester.

> Done milestones (Phase 1, Phase 2, etc.) are summarized in [`COMPLETED.md`](COMPLETED.md). What follows is forward-looking only.

---

## Guiding principle: improve before audience-reducing changes

Before any change that would reduce the project's reach &mdash; TFM bumps, modern-.NET moves, Visio version baseline shifts &mdash; we prioritize work that makes the *current* audience better served: documentation, convenience features, code quality, testing, maintainability, and addressing user feedback. The reasoning: if anyone later wants to pick up maintenance of an older-Visio or older-.NET fork, they should start from a finished, well-documented, well-tested baseline rather than inheriting a half-renovated codebase.

This principle drives the semester sequencing below. **2026 (CY26Q2 through CY26Q4)** is dedicated to audience-preserving improvements. **2027 onward (CY27Q1+)** is when audience-reducing modernization begins, after the Windows 10 LTSB 2016 mainstream-support sunset (2026-10-13) and after the 2026 priority work is mature.

---

## Semester schedule

### CY26Q2 &mdash; April&ndash;June 2026 *(current)*
[GitHub milestone](https://github.com/saveenr/VisioAutomation/milestone/2) &mdash; due 2026-06-30

**Theme:** Kick off audience-preserving improvements. Schedule the cmdlet-ergonomics review, start the portfolio audit, sweep older issues, formalize the test-suite design decision.

| Item | Themed milestone | Issue |
|---|---|---|
| Scoping review (May 2026): VisioBot3000 + PSVA borrowed-ideas backlog | D | [#149](https://github.com/saveenr/VisioAutomation/issues/149) |
| PSVA cmdlet-surface audit (feeds the May review) | D | [#150](https://github.com/saveenr/VisioAutomation/issues/150) |
| Q2 2026 issue triage pass (#80, #82, #102, #105, #117) | Triage | [#151](https://github.com/saveenr/VisioAutomation/issues/151) |
| Visio repo portfolio audit Phase H1 (centralized index + per-repo READMEs) | H | [#152](https://github.com/saveenr/VisioAutomation/issues/152) |
| Tests-need-Visio design-decision write-up | F | [#153](https://github.com/saveenr/VisioAutomation/issues/153) |
| Version compatibility tables on user-facing gitbooks | E | [#161](https://github.com/saveenr/VisioAutomation/issues/161) |

### CY26Q3 &mdash; July&ndash;September 2026
[GitHub milestone](https://github.com/saveenr/VisioAutomation/milestone/3) &mdash; due 2026-09-30

**Theme:** Execute on the May 2026 cmdlet-ergonomics shortlist; close the docs and test decisions that are blocking later work; ship the existing docs-coverage backlog.

| Item | Themed milestone | Issue |
|---|---|---|
| `VisioScripting.Client` undocumented in gitbook | E | [#131](https://github.com/saveenr/VisioAutomation/issues/131) |
| `VisioAutomation.Models` undocumented (DOM + Layouts) | E | [#132](https://github.com/saveenr/VisioAutomation/issues/132) |
| Add troubleshooting page to .NET gitbook | E | [#133](https://github.com/saveenr/VisioAutomation/issues/133) |
| Audit: test coverage gaps on the public API surface | F | [#154](https://github.com/saveenr/VisioAutomation/issues/154) |
| Discussion: doc-sample-as-test linkage mechanism | E | [#155](https://github.com/saveenr/VisioAutomation/issues/155) |
| Decide: VisioScripting public-API status | E | [#156](https://github.com/saveenr/VisioAutomation/issues/156) |
| Decide: long-term docs location | E | [#157](https://github.com/saveenr/VisioAutomation/issues/157) |
| (Cmdlet ergonomics implementation work from the May shortlist) | D | (issues filed in CY26Q2 from the [#149](https://github.com/saveenr/VisioAutomation/issues/149) outcome) |

### CY26Q4 &mdash; October&ndash;December 2026
[GitHub milestone](https://github.com/saveenr/VisioAutomation/milestone/1) &mdash; due 2026-12-31

**Theme:** Identity transition push. Move GitHub repo and gitbook spaces to SevenPens-owned hosting; rewrite in-repo URLs once the destinations are stable.

| Item | Themed milestone | Issue |
|---|---|---|
| Migrate GitHub repo to SevenPens-owned account | A | [#146](https://github.com/saveenr/VisioAutomation/issues/146) |
| Migrate gitbook spaces to SevenPens-owned hosting | A | [#147](https://github.com/saveenr/VisioAutomation/issues/147) |
| Phase 5b: in-repo URL rewrite (single commit after #146 + #147 land) | A | (no issue; one-shot commit) |
| Retire unused `VisioAutomation` PSGallery co-owner account | A | [#148](https://github.com/saveenr/VisioAutomation/issues/148) (ride-along) |

### CY27Q1 &mdash; January&ndash;March 2027
[GitHub milestone](https://github.com/saveenr/VisioAutomation/milestone/4) &mdash; due 2027-03-31

**Theme:** Modernization unlock (gated on Windows 10 LTSB 2016 mainstream-support sunset 2026-10-13). TFM bump, VS 2026, manifest deprecation cleanups. Plus the spike-and-audit work that informs 2027's deeper changes.

| Item | Themed milestone | Issue |
|---|---|---|
| TFM bump shipping-libs 4.5.2 &rarr; 4.7.2 | B | (no issue yet; file when starting) |
| Move dev environment to VS 2026 | B | (no issue yet) |
| `Visio.psd1` deprecation cleanups (`ModuleToProcess`/`PowerShellVersion`) | B | (no issue yet) |
| Re-evaluate version-policy decision | B | (no issue yet) |
| Switch module-release builds Debug &rarr; Release | B | (no issue yet) |
| Spike: NetOffice / NetOfficeFw as a Visio PIA replacement | C | [#158](https://github.com/saveenr/VisioAutomation/issues/158) |
| Audit: identify Visio-2010-only paths in the codebase | G | [#159](https://github.com/saveenr/VisioAutomation/issues/159) |

### CY27Q2 &mdash; April&ndash;June 2027
[GitHub milestone](https://github.com/saveenr/VisioAutomation/milestone/5) &mdash; due 2027-06-30

**Theme:** Modern .NET multi-target work begins; Visio 2013 baseline decisions. Both informed by the CY27Q1 spike and audit.

| Item | Themed milestone | Issue |
|---|---|---|
| C# 14 / .NET 10 multi-target migration (informed by [#158](https://github.com/saveenr/VisioAutomation/issues/158)) | C | (no issue yet; file from #158 outcome) |
| Visio PIA replacement decision (informed by [#158](https://github.com/saveenr/VisioAutomation/issues/158)) | C | (folded into #158's deliverable) |
| Decide: Visio 2013 baseline branding | G | [#160](https://github.com/saveenr/VisioAutomation/issues/160) |
| Visio 2013 baseline migration implementation (gated on #160) | G | (no issue yet) |

### Beyond CY27Q2

Items pending but not yet semester-assigned:
- Documentation work past the CY26Q3 burst &mdash; treated as continuous through CY26Q4 and into 2027 as items come up. New docs items get assigned to the semester they're picked up in.
- Visio repo portfolio audit Phase H2 (per-repo retire / merge / maintain decisions). Opportunistic through CY26Q4 and CY27 as the picture clarifies.

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

**Items:** see CY27Q1 schedule above. Detail in [`futures/build-and-code.md`](futures/build-and-code.md) and [`futures/releases.md`](futures/releases.md).

### Milestone C &mdash; Modern .NET (multi-target)
**Theme:** Add `net10.0-windows` (or current .NET 10 TFM) alongside `net48`. Decide whether to replace the Visio PIA with NetOffice / NetOfficeFw.

**User impact:** Modern C# features (extension members, etc.) and better .NET 10 perf become available. Existing `net48` consumers unaffected (multi-target). Windows PowerShell 5.1 users of VisioPS get the `net48` build; PowerShell 7+ users get the .NET 10 build.

**Semesters:** CY27Q1 (spike) &rarr; CY27Q2 (implementation begins). Sequenced after Milestone B.

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
- Tier 3 .NET-side coverage (`VisioAutomation.Models` project) &mdash; covered by [#132](https://github.com/saveenr/VisioAutomation/issues/132).
- Restructure user-docs repos &mdash; gated on the docs-location decision ([#157](https://github.com/saveenr/VisioAutomation/issues/157)).
- Five smaller gitbook items &mdash; see [`futures/docs.md`](futures/docs.md).
- Keep CHANGELOGs current (process item, not a one-shot).
- **Visio version &harr; PIA mapping reference page** &mdash; new doc page explaining the version-number table (Visio 2010 = 14, 2013 = 15, ...) and where to get each PIA. Couples to the [`Visio-PIAs`](https://github.com/saveenr/Visio-PIAs) sibling repo.

### Milestone F &mdash; Test infrastructure
**Theme:** Modernize the test suite and address the live-Visio dependency that currently keeps tests off CI.

**User impact:** More reliable releases (catch regressions in CI rather than at publish time); contributor onboarding without a Visio install for the test buckets that don't actually need one.

**Semesters:** CY26Q2 (design-decision write-up) &rarr; CY26Q3 (coverage audit) &rarr; CY26Q4+ (implementation as bandwidth allows).

**Items:** see CY26Q2 and CY26Q3 schedules above. Detail in [`futures/build-and-code.md`](futures/build-and-code.md) and [`futures/tests.md`](futures/tests.md). The "tests need a live Visio" rule that gates CI today is now formalized in [`decisions/tests-need-visio.md`](decisions/tests-need-visio.md). The no-Visio test split has *de facto* started: [`VisioPS_Manifest_Tests`](../VisioAutomation_2010/VTest.PowerShell/VisioPS_Manifest_Tests.cs) and [`XmlErrorLogTests`](../VisioAutomation_2010/VTest/Core/Application/XmlErrorLogTests.cs) are existing examples of the no-Visio bucket; the planned tagging work to turn the implicit split explicit is part of this milestone.

### Milestone G &mdash; Visio 2013 baseline migration
**Theme:** Move the codebase's baseline Visio version from 2010 to 2013.

**User impact:** *Audience-reducing.* Consumers still on Visio 2010 (released 2010, mainstream support ended 2015, extended support ended 2020) would no longer see new releases. Visio 2013+ users get access to capabilities the 2013-era Visio API added.

**Semesters:** CY27Q1 (audit) &rarr; CY27Q2 (branding decision &rarr; implementation).

**Items:** see CY27Q1 and CY27Q2 schedules above.

### Milestone H &mdash; Visio repo portfolio audit
**Theme:** Take stock of the broader Visio-related repo portfolio. Make each repo's status visible; decide what to retire / merge / maintain.

**User impact:** Visitors to any of the 9 sibling repos can tell at a glance whether it's active. Reduces "is this still maintained?" friction. Cleanup of dormant repos shrinks reader surface area.

**Semesters:** CY26Q2 (Phase H1: index + per-repo READMEs). CY26Q3+ for Phase H2 (per-repo retire/merge/maintain decisions, opportunistic).

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

---

## Triage backlog (handled in [#151](https://github.com/saveenr/VisioAutomation/issues/151), CY26Q2)

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
