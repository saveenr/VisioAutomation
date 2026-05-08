# Tests directly drive a real Visio instance via COM

**Status:** Accepted (codified 2026-05-07; choice itself dates to project inception)

## Context

The `VisioAutomation` library is a thin layer over Microsoft Visio's COM API. Its entire job is to drive Visio. The test infrastructure has to choose between two approaches:

1. **Mock the Visio COM surface.** Build a fake `Microsoft.Office.Interop.Visio.Application` (and the rest of the COM surface) that responds to API calls in-memory.
2. **Drive a real Visio instance.** Every test launches Visio (or shares a per-testhost singleton) and asserts against actual ShapeSheet behavior.

## Decision

Approach 2, exclusively. There is no mock or fake Visio. Every test in `VTest`, `VTest.Models`, `VTest.Scripting`, and `VTest.PowerShell` either drives a real `IVisio.Application` or tests something that doesn't touch Visio at all (see *No-Visio bucket* below).

## Why

- **A mock has to encode Visio's actual behavior, including its undocumented quirks.** When the library produces output that Visio interprets in a specific way, a mock that doesn't match Visio's interpretation produces tests that are confidently wrong. Tests against the mock then prove that the mock matches the mock, not that the library matches Visio.
- **The library's whole purpose is the COM-Visio bridge.** Mocking the COM surface tests everything *except* the part that actually matters.
- **Empirical confirmation.** Bugs caught during 2026 illustrate the kind of failure that only a real Visio surfaces:
  - `OrgChartStyling.Visio2013Template = "orgch_u.vstx"` (commit `da9bba0a`) &mdash; Visio 2013 silently changed binary `.vst` templates to XML `.vstx` and modern Visio installs only ship the new form. A mock would not have known.
  - `MsaglRenderer` Size + Cells precedence (commit `805f90df`, closed [#82](https://github.com/saveenr/VisioAutomation/issues/82)) &mdash; the bug only manifested as wrong rendered Width/Height on a real Visio shape after the DOM emitted formulas; a mock would have happily reported back whatever Width/Height our DOM said it set.
- **Industry alignment.** The .NET community has broadly moved away from heavy COM mocking. The CONTRIBUTING.md "no mocks" rule predates this shift but aligns with it.

## Consequences

### Operational

- **Tests cannot run on a machine without Visio installed.** This is binding for both contributors and CI.
- **CI today is build-only.** [`.github/workflows/build.yml`](../../.github/workflows/build.yml) restores and builds; it does not run tests. Running tests in CI requires a self-hosted Windows runner with Visio installed (tracked in [`futures/build-and-code.md`](../futures/build-and-code.md#run-tests-in-ci) and Milestone F).
- **Tests run sequentially, one Visio instance per testhost.** Visio doesn't tolerate concurrent COM clients well. The per-testhost singleton ([`Framework.VTestAppRef`](../../VisioAutomation_2010/VTest/Framework/VTestAppRef.cs)) and `[AssemblyCleanup]` orphan-prevention machinery exist to make sequential testing economical despite Visio's cold-start cost. See [`docs/TESTING.md`](../TESTING.md) for the operational details.
- **Test failures can be environmental.** "Works on my machine" failures are slightly more common than in pure-logic test suites because Visio version, installed templates / stencils, and locale settings all participate in test outcomes.

### Strategic

- **No `dotnet test` that runs anywhere.** New contributors hit this immediately when they try the suite without Visio. [`CONTRIBUTING.md`](../../CONTRIBUTING.md) flags the dependency upfront.
- **Future CI cost.** A self-hosted Windows runner with a Visio license is the only viable path for running the full suite in CI. Until then, the no-Visio bucket below is the only piece that can run on `windows-latest`.

## No-Visio bucket (de facto split)

A small subset of tests does *not* need Visio because it tests pure file-I/O or metadata, not COM behavior. Two extant examples as of 2026-05-07:

- [`ManifestTests`](../../VisioAutomation_2010/VTest.PowerShell/ManifestTests.cs) &mdash; `CmdletsToExport` drift check between the compiled assembly and the `Visio.psd1` manifest.
- [`XmlErrorLogTests`](../../VisioAutomation_2010/VTest/Core/Application/XmlErrorLogTests.cs) &mdash; XML parsing of pre-captured Visio error logs.

These tests don't reference `IVisio.Application`, don't go through `Framework.VTest.GetVisioApplication()`, and don't trip the singleton lifecycle. They run identically with or without a Visio install.

This split is currently *implicit* &mdash; you can tell which bucket a test is in by reading it, but nothing in the project structure marks it. The intentional version of this split is the foundation of the eventual CI-tests-without-Visio plan: tag the no-Visio bucket explicitly (test category, separate project, or naming convention &mdash; TBD) so CI can run that bucket on `windows-latest` without a Visio license. That tagging work is part of Milestone F.

## Cross-references

- [`docs/TESTING.md`](../TESTING.md) &mdash; operational details (singleton, `[AssemblyCleanup]`, MSTEST0030, etc.)
- [`CONTRIBUTING.md`](../../CONTRIBUTING.md) &mdash; the contributor-facing one-liner, pointing here.
- [`docs/futures/build-and-code.md`](../futures/build-and-code.md#run-tests-in-ci) &mdash; *Run tests in CI* backlog item.
- [`docs/futures/tests.md`](../futures/tests.md) &mdash; test backlog (no longer carries this decision; defers here).
- [`docs/MILESTONES.md`](../MILESTONES.md) &mdash; Milestone F (Test infrastructure).

## Reconsider when

- **Visio gets a headless mode.** None in sight; Microsoft has not signaled this.
- **A vendor ships a Visio-equivalent COM mock with maintained behavior parity.** No candidate today.
- **The library's scope expands beyond COM-driving** (e.g., adding pure-data analysis features that don't touch Visio). New code paths could ship with non-Visio tests without changing this decision for the existing surface; the no-Visio bucket would simply grow.
