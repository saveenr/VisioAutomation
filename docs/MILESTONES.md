# MILESTONES.md &mdash; forward-looking work plan

Themed groupings of the [`FUTURES.md`](FUTURES.md) backlog into milestones with proposed target windows. Sits between [`ROADMAP.md`](ROADMAP.md) (high-level phase plan) and [`FUTURES.md`](FUTURES.md) (backlog index, no scheduling).

Targets are *proposals*. Items inside a milestone can land independently &mdash; the milestone is a thematic grouping plus rough timing, not a "ship together" commitment. Where a milestone has a hard external gate (e.g., a Windows support sunset), that's flagged.

> Done milestones (Phase 1, Phase 2, etc.) are summarized in [`COMPLETED.md`](COMPLETED.md). What follows is forward-looking only.

---

## Guiding principle: improve before audience-reducing changes

Before any change that would reduce the project's reach &mdash; TFM bumps, modern-.NET moves, Visio version baseline shifts &mdash; we prioritize work that makes the *current* audience better served: documentation, convenience features, code quality, testing, maintainability, and addressing user feedback. The reasoning: if anyone later wants to pick up maintenance of an older-Visio or older-.NET fork, they should start from a finished, well-documented, well-tested baseline rather than inheriting a half-renovated codebase.

This principle drives the milestone sequencing below. **Audience-preserving milestones get the rest of 2026** &mdash; A (Dec 2026), D, E, F, H. **Audience-reducing milestones (B, G, C) are sequenced for 2027 onward**, behind both the LTSB 2016 mainstream-support sunset and the 2026 priority work.

| Window | Priority milestones |
|---|---|
| 2026 (now &rarr; December) | A &mdash; Identity transition (Dec target). D, E, F, H &mdash; audience-preserving improvements. |
| Q1 2027 | B &mdash; Modernization unlock (gated on Windows 10 LTSB 2016 sunset, 2026-10-13). |
| 2027 onward | G &mdash; Visio 2013 baseline migration. C &mdash; Modern .NET multi-target. |

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

## Milestone D &mdash; Cmdlet ergonomics
**Target:** Scoping review **May 2026**; implementation through the rest of 2026 as bandwidth allows (no hard ship date)

**Goal:** Adopt the good ideas from sibling community PowerShell-for-Visio projects to lift VisioPS authoring quality.

**User impact:** Friendlier authoring for VisioPS users &mdash; nickname registry, block-style nesting for containers, bulk shape operations, custom-property authoring without manual `EncodeValues()`. Limited to the items the May 2026 scoping review picks out; everything else stays in the futures backlog as "considered, not pursued this year".

**Plan:**
1. **May 2026 scoping review** (~half day): walk the full borrowed-ideas backlog ([VisioBot3000](https://github.com/cofonseca/VisioBot3000), [PSVA](https://github.com/dotps1/PSVA), the `EncodeValues` simplification) and produce a shortlist with rough scope and ordering for 2026. Output goes back into [`futures/build-and-code.md`](futures/build-and-code.md) and into this milestone as concrete items.
2. **Implementation through the rest of 2026** as bandwidth allows. Items not picked up by year-end stay in the backlog without being lost.

**Candidate items for the review** (full detail in [`futures/build-and-code.md`](futures/build-and-code.md)):
- VisioBot3000 ideas: nickname registry, dynamic functions per registered shape, block-style syntax, relative-position cursor.
- PSVA ideas: bulk shape distribution, pipeline-friendly bulk connectors, side-and-alignment shape decoration, layer cmdlets.
- Make `CustomPropertyCells` values not require manual `EncodeValues()`.
- Move `LinqExtensions` out of `Internal/` if it has public-API utility.

---

## Milestone E &mdash; Documentation completeness
**Target:** Continuous through 2026 (per the guiding principle's emphasis on audience-preserving work; no single ship event)

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
- **Visio version &harr; PIA mapping reference page.** "Which Visio version uses which Primary Interop Assembly?" and "where do I get them?" come up regularly and confuse newcomers because Microsoft's version numbers (Visio 2010 = 14, 2013 = 15, 2016 = 16, ...) don't match the marketing names. A single canonical reference page on the gitbook would defuse this; should cover the version-number table, where to get each PIA (the [`saveenr/Visio-PIAs`](https://github.com/saveenr/Visio-PIAs) repo handles this for the older versions; modern PIAs are NuGet packages), and which the codebase currently bundles.
- **Doc-sample-as-test linkage** *(discussion-needed first; mechanism open)*: a long-standing fantasy that every code sample in the public docs corresponds to a unit test, linked by name in a sample comment, so samples can never bit-rot. The mechanism is open &mdash; could be as light as a naming convention (e.g., sample comment `// see VisioPS_DrawRectangleWithText`), or as heavy as a sample-extraction tool that lifts code blocks from gitbook markdown into test files. Worth a focused discussion before committing to a particular approach.

---

## Milestone F &mdash; Test infrastructure
**Target:** 2026 (per the guiding principle's "improve-before-audience-reducing" framing); no single ship event

**Goal:** Modernize the test suite and address the live-Visio dependency that currently keeps tests off CI.

**User impact:** More reliable releases (catch regressions in CI rather than at publish time); contributor onboarding without a Visio install for the test buckets that don't actually need one.

**Items:**
- Run tests in CI: requires figuring out the live-Visio dependency. Likely path: split tests into "needs Visio" / "doesn't need Visio" buckets and run the latter on GitHub Actions ([`futures/build-and-code.md`](futures/build-and-code.md), [`futures/tests.md`](futures/tests.md)). The recently-added [`VisioPS_Manifest_Tests`](../VisioAutomation_2010/VTest.PowerShell/VisioPS_Manifest_Tests.cs) and [`XmlErrorLogTests`](../VisioAutomation_2010/VTest/Core/Application/XmlErrorLogTests.cs) are existing examples of the no-Visio bucket; the split has *de facto* started.
- Address test coverage gaps ([`futures/tests.md`](futures/tests.md)).
- Evaluate modern testing-stack options ([`futures/tests.md`](futures/tests.md)).
- Cross-ref: the doc-sample-as-test work in Milestone E uses the same test-suite mechanics; coordinate if both move at once.

---

## Milestone H &mdash; Visio repo portfolio audit
**Target:** Centralized index + README status updates in 2026 (cheap); deeper decisions (retire / merge) opportunistic through 2027

**Goal:** Take stock of the broader Visio-related repo portfolio. Make each repo's status visible to visitors, decide what to retire, what to merge, what to keep maintaining.

**User impact:** Visitors to any of the repos can tell at a glance whether it's active, archived, or superseded &mdash; reduces "is this still maintained?" friction. Cleanup of dormant repos shrinks the surface area readers have to navigate.

**Repos in scope:**
- [`saveenr/VisioAutomation`](https://github.com/saveenr/VisioAutomation) &mdash; this repo. Active.
- [`saveenr/VisioAutomation2007`](https://github.com/saveenr/VisioAutomation2007) &mdash; older Visio 2007 baseline.
- [`saveenr/VisioAutomation.VDX`](https://github.com/saveenr/VisioAutomation.VDX) &mdash; VDX format work.
- [`saveenr/Visio-PIAs`](https://github.com/saveenr/Visio-PIAs) &mdash; distribution of Visio Primary Interop Assemblies (relevant to the Milestone E "PIA mapping" doc work).
- [`saveenr/visio-templates`](https://github.com/saveenr/visio-templates) &mdash; Visio templates.
- [`saveenr/visio-reference`](https://github.com/saveenr/visio-reference) &mdash; reference data.
- [`saveenr/Visio-Power-Tools`](https://github.com/saveenr/Visio-Power-Tools) &mdash; Visio Power Tools.
- [`saveenr/Visio-Export-Pages-To-Docs`](https://github.com/saveenr/Visio-Export-Pages-To-Docs) &mdash; export utility.
- [`saveenr/Visio-Font-Compare`](https://github.com/saveenr/Visio-Font-Compare) &mdash; font comparison.
- [`saveenr/Visio-Code-Samples`](https://github.com/saveenr/Visio-Code-Samples) &mdash; code samples.

**Items:**
- **Phase H1 (2026, cheap):** centralized index. A single page lists every repo above with current status (active / archived / paused / superseded), a one-line description, and a link. Could live in this repo's `readme.md`, or as a new `docs/RELATED-REPOS.md`, or as a pinned-repo arrangement on the GitHub profile.
- **Phase H1 (2026, cheap):** per-repo README updates. Each repo's `README.md` clearly states `Status: active / archived / paused / superseded by X` plus a one-line description and any pointers to successor projects. No code changes; just docs.
- **Phase H2 (opportunistic, 2026&ndash;2027):** per-repo decisions. For each, decide: retire (archive on GitHub), maintain (no further changes), merge into VisioAutomation, or merge across siblings. Repos that should merge get a separate transition plan per merge.
- **Cross-ref:** when the GitHub repo move (Axis 5a-1, [#146](https://github.com/saveenr/VisioAutomation/issues/146)) happens for VisioAutomation, the same hosting-URL question arises for each sibling repo. Coordinated transfer (all at once vs. staged) should be decided as part of Phase H2.

---

## Milestone B &mdash; Modernization unlock
**Target:** Q1 2027 *(gated: Windows 10 LTSB 2016 mainstream-support sunset is 2026-10-13)*

**Goal:** Bump shipping-lib TFMs from .NET Framework 4.5.2 to 4.7.2, move dev to VS 2026, and clear the longstanding deprecation / version-hygiene cruft.

**User impact:** *Audience-reducing.* Minimum supported .NET Framework rises from 4.5.2 to 4.7.2 &mdash; affects only consumers on pre-1803 Windows, which leave mainstream support October 2026. Modern IDE tooling, faster builds, cleaner manifest metadata. Sequenced after the 2026 audience-preserving work per the guiding principle.

**Items:**
- TFM bump shipping-libs 4.5.2 &rarr; 4.7.2 (gated on the LTSB 2016 sunset; see the [enterprise-compat memory](../../.claude/projects/C--Users-savee-Documents-GitHub-VisioAutomation/memory/enterprise_compat_ltsb2016.md)).
- Move dev environment to Visual Studio 2026 ([`futures/build-and-code.md`](futures/build-and-code.md)).
- `Visio.psd1` deprecation cleanups: `ModuleToProcess` &rarr; `RootModule`, `PowerShellVersion '2.0'` &rarr; `'5.1'` ([`futures/releases.md`](futures/releases.md)). Customer impact already analyzed as zero.
- Re-evaluate version-policy decision (currently "stay divergent"; the PS-5.1 manifest bump above is the agreed forcing function).
- Switch module-release builds from Debug to Release ([`futures/releases.md`](futures/releases.md)).

---

## Milestone G &mdash; Visio 2013 baseline migration
**Target:** TBD, 2027+ (audience-reducing; sequenced after the 2026 priorities and probably after Milestone B)

**Goal:** Move the codebase's baseline Visio version from 2010 to 2013.

**User impact:** *Audience-reducing.* Consumers still on Visio 2010 (released 2010, mainstream support ended 2015, extended support ended 2020) would no longer see new releases &mdash; though anyone actively maintaining a Visio-2010 path can pick up the last 2010-baselined release per the guiding principle's intent. Visio 2013+ users get access to capabilities the 2013-era Visio API added (improved theming, the modern `.vsdx` XML format as the default save target, etc.).

**Items:**
- Decide branding: continue `VisioAutomation2010` indefinitely with newer Visio APIs added on top, OR cut a parallel `VisioAutomation2013` package, OR rename in-place. Ties into the version-policy decision (deferred from Milestone B).
- Update bundled PIA from Visio 2010 to Visio 2013. Cross-refs the Visio PIA mapping doc (Milestone E) and the [`Visio-PIAs`](https://github.com/saveenr/Visio-PIAs) sibling repo (Milestone H).
- Audit code for Visio-2010-only paths that can be removed. Example precedent: the orgchart `.vst` template fallback in [`OrgChartStyling.cs`](../VisioAutomation_2010/VisioAutomation/OrgChart/OrgChartStyling.cs), which was patched May 2026 to handle 2013+ defaulting to `.vstx`. Similar version-guarded paths likely exist elsewhere.
- Documentation refresh: clearly mark which Visio version is the baseline, both in the readme and in user-facing docs.

---

## Milestone C &mdash; Modern .NET (multi-target)
**Target:** Q2&ndash;Q3 2027 (after Milestone B; possibly overlaps with Milestone G)

**Goal:** Add `net10.0-windows` (or whichever .NET 10 TFM is current) alongside `net48`. Decide whether to replace the Visio PIA with NetOffice / NetOfficeFw.

**User impact:** Modern C# features (extension members, etc.) and better .NET 10 perf become available. Existing `net48` consumers unaffected (multi-target). Windows PowerShell 5.1 users of VisioPS get the `net48` build; PowerShell 7+ users get the .NET 10 build.

**Items:**
- NetOffice / NetOfficeFw spike (~1 day, go/no-go memo on whether it covers the codebase's COM surface and ships modern-.NET TFMs) ([`futures/build-and-code.md`](futures/build-and-code.md)).
- Move to C# 14 / .NET 10 multi-target ([`futures/build-and-code.md`](futures/build-and-code.md)).
- Visio PIA replacement decision (informed by the NetOffice spike). May coordinate with Milestone G's PIA-version bump.

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
- Add new milestones as the long-term picture clarifies. Eight is a working number, not a fixed bucket count.
- The targets here are not commitments to consumers &mdash; release notes and CHANGELOGs are the authoritative "what's shipped when" record. This doc is internal planning.
- The guiding principle (improve before audience-reducing changes) governs sequencing decisions. If a future milestone bumps a 2027+ item earlier, the principle should be revisited at the same time, not silently overruled.
