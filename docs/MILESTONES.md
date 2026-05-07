# MILESTONES.md &mdash; forward-looking work plan

Themed groupings of the [`FUTURES.md`](FUTURES.md) backlog into milestones with proposed target windows. Sits between [`ROADMAP.md`](ROADMAP.md) (high-level phase plan) and [`FUTURES.md`](FUTURES.md) (backlog index, no scheduling).

Targets are *proposals*. Items inside a milestone can land independently &mdash; the milestone is a thematic grouping plus rough timing, not a "ship together" commitment. Where a milestone has a hard external gate (e.g., a Windows support sunset), that's flagged.

> Done milestones (Phase 1, Phase 2, etc.) are summarized in [`COMPLETED.md`](COMPLETED.md). What follows is forward-looking only.

---

## Milestone A &mdash; Identity transition complete
**Target:** December 2026 ([GH milestone `2026-12`](https://github.com/saveenr/VisioAutomation/milestone/1))

**Goal:** Finish the dev-team identity transition from `Saveen` / `saveenr` to `SevenPens`.

**User impact:** Old hosting URLs (`github.com/saveenr/...`, `saveenr.gitbook.io/...`) keep resolving via redirects; new canonical URLs live under `SevenPens`-owned hosting. No broken bookmarks, no consumer disruption.

**Items:**
- [#146](https://github.com/saveenr/VisioAutomation/issues/146) &mdash; Migrate GitHub repo to SevenPens-owned account.
- [#147](https://github.com/saveenr/VisioAutomation/issues/147) &mdash; Migrate gitbook spaces to SevenPens-owned hosting.
- Phase 5b in-repo URL rewrite &mdash; single mechanical commit after #146 + #147 land. See [`futures/identity.md`](futures/identity.md). No issue (will be a one-shot commit).
- [#148](https://github.com/saveenr/VisioAutomation/issues/148) &mdash; Retire unused `VisioAutomation` PSGallery co-owner account (ride-along; unscheduled but fits naturally here).

---

## Milestone B &mdash; Modernization unlock
**Target:** Q1 2027 *(gated: Windows 10 LTSB 2016 mainstream-support sunset is 2026-10-13)*

**Goal:** Bump shipping-lib TFMs from .NET Framework 4.5.2 to 4.7.2, move dev to VS 2026, and clear the longstanding deprecation / version-hygiene cruft.

**User impact:** Minimum supported .NET Framework rises from 4.5.2 to 4.7.2 &mdash; affects only consumers on pre-1803 Windows, which leave mainstream support October 2026. Modern IDE tooling, faster builds, cleaner manifest metadata.

**Items:**
- TFM bump shipping-libs 4.5.2 &rarr; 4.7.2 (gated on the LTSB 2016 sunset; see the [enterprise-compat memory](../../.claude/projects/C--Users-savee-Documents-GitHub-VisioAutomation/memory/enterprise_compat_ltsb2016.md)).
- Move dev environment to Visual Studio 2026 ([`futures/build-and-code.md`](futures/build-and-code.md)).
- `Visio.psd1` deprecation cleanups: `ModuleToProcess` &rarr; `RootModule`, `PowerShellVersion '2.0'` &rarr; `'5.1'` ([`futures/releases.md`](futures/releases.md)). Customer impact already analyzed as zero.
- Re-evaluate version-policy decision (currently "stay divergent"; the PS-5.1 manifest bump above is the agreed forcing function).
- Switch module-release builds from Debug to Release ([`futures/releases.md`](futures/releases.md)).

---

## Milestone C &mdash; Modern .NET (multi-target)
**Target:** Q2&ndash;Q3 2027 (after Milestone B)

**Goal:** Add `net10.0-windows` (or whichever .NET 10 TFM is current) alongside `net48`. Decide whether to replace the Visio 2010 PIA with NetOffice / NetOfficeFw.

**User impact:** Modern C# features (extension members, etc.) and better .NET 10 perf become available. Existing `net48` consumers unaffected (multi-target). Windows PowerShell 5.1 users of VisioPS get the `net48` build; PowerShell 7+ users get the .NET 10 build.

**Items:**
- NetOffice / NetOfficeFw spike (~1 day, go/no-go memo on whether it covers the codebase's COM surface and ships modern-.NET TFMs) ([`futures/build-and-code.md`](futures/build-and-code.md)).
- Move to C# 14 / .NET 10 multi-target ([`futures/build-and-code.md`](futures/build-and-code.md)).
- Visio 2010 PIA replacement decision (informed by the NetOffice spike).

---

## Milestone D &mdash; Cmdlet ergonomics
**Target:** Opportunistic (no hard date; bandwidth-driven)

**Goal:** Adopt good ideas from sibling community PowerShell-for-Visio projects to lift VisioPS authoring quality.

**User impact:** Friendlier authoring for VisioPS users &mdash; nickname registry, block-style nesting for containers, bulk shape operations, custom-property authoring without manual `EncodeValues()`.

**Items:**
- Borrow ideas from VisioBot3000: nickname registry, dynamic functions per registered shape, block-style syntax, relative-position cursor ([`futures/build-and-code.md`](futures/build-and-code.md)).
- Borrow ideas from PSVA: bulk shape distribution, pipeline-friendly bulk connectors, side-and-alignment shape decoration, layer cmdlets ([`futures/build-and-code.md`](futures/build-and-code.md)).
- Make `CustomPropertyCells` values not require manual `EncodeValues()` ([`futures/build-and-code.md`](futures/build-and-code.md)).
- Move `LinqExtensions` out of `Internal/` if it has public-API utility ([`futures/build-and-code.md`](futures/build-and-code.md)).

---

## Milestone E &mdash; Documentation completeness
**Target:** Continuous (land items as they come up; no single ship event)

**Goal:** Close the documentation gaps on the user-facing gitbooks and the in-repo dev guides.

**User impact:** Easier onboarding for new contributors and consumers; fewer "had to read the source" moments. Reference docs cover the surface that users actually program against.

**Items:**
- [#131](https://github.com/saveenr/VisioAutomation/issues/131) &mdash; `VisioScripting.Client` undocumented in gitbook.
- [#132](https://github.com/saveenr/VisioAutomation/issues/132) &mdash; `VisioAutomation.Models` undocumented (DOM + Layouts).
- [#133](https://github.com/saveenr/VisioAutomation/issues/133) &mdash; Add troubleshooting page to .NET gitbook.
- Tier 3 .NET-side coverage (`VisioAutomation.Models` project) ([`futures/docs.md`](futures/docs.md)).
- Decide whether to document `VisioScripting` as a public API ([`futures/docs.md`](futures/docs.md)).
- Decide where docs live long-term &mdash; in-repo vs. external gitbook split ([`futures/docs.md`](futures/docs.md)).
- Restructure user-docs repos ([`futures/docs.md`](futures/docs.md)).
- Five smaller gitbook items ([`futures/docs.md`](futures/docs.md)).
- Keep CHANGELOGs current (process item, not a one-shot) ([`futures/docs.md`](futures/docs.md)).

---

## Milestone F &mdash; Test infrastructure
**Target:** Opportunistic (no hard date)

**Goal:** Modernize the test suite and address the live-Visio dependency that currently keeps tests off CI.

**User impact:** More reliable releases (catch regressions in CI rather than at publish time); contributor onboarding without a Visio install for the test buckets that don't actually need one.

**Items:**
- Run tests in CI: requires figuring out the live-Visio dependency. Likely path: split tests into "needs Visio" / "doesn't need Visio" buckets and run the latter on GitHub Actions ([`futures/build-and-code.md`](futures/build-and-code.md), [`futures/tests.md`](futures/tests.md)).
- Address test coverage gaps ([`futures/tests.md`](futures/tests.md)).
- Evaluate modern testing-stack options ([`futures/tests.md`](futures/tests.md)).

---

## Triage backlog (not in any milestone)

Older issues that deserve a closer look before being slotted into a milestone or closed. Listed here so they're not lost.

- [#80](https://github.com/saveenr/VisioAutomation/issues/80) &mdash; New logo for Visio Automation. Cosmetic; could ride into Milestone A or E.
- [#82](https://github.com/saveenr/VisioAutomation/issues/82), [#102](https://github.com/saveenr/VisioAutomation/issues/102) &mdash; older user-help questions. Probably close-as-answered or convert to gitbook how-to entries (Milestone E).
- [#105](https://github.com/saveenr/VisioAutomation/issues/105) &mdash; directed-graph layout/direction umbrella. Sub-items shipped May 2026; the umbrella issue may be ready to close.
- [#117](https://github.com/saveenr/VisioAutomation/issues/117) &mdash; custom-properties on directed graph. Reporter pinged with the fix from #144 (typed setters); awaiting their confirmation to close.

---

## Updating this doc

- When a milestone item lands, move it from this doc to [`COMPLETED.md`](COMPLETED.md) (or strike it through here if the milestone is mid-flight).
- When a proposed target slips or changes, update the target line at the top of the milestone.
- Add new milestones as the long-term picture clarifies. Six is a working number, not a fixed bucket count.
- The targets here are not commitments to consumers &mdash; release notes and CHANGELOGs are the authoritative "what's shipped when" record. This doc is internal planning.
