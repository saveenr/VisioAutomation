
# Doc samples are checked by compiling, not by running

**Status:** Accepted (2026-11-20; see [#155](https://github.com/saveenr/VisioAutomation/issues/155) for the discussion)

## Context

[#155](https://github.com/saveenr/VisioAutomation/issues/155) asked how the C# code samples in the public .NET docs could be kept from going stale. A sample can go bad in three ways: it stops compiling (an API was renamed or removed), the docs text drifts from the verified code, or it compiles but behaves wrongly. The issue listed four mechanisms: a naming convention, a test-name registry, extracting samples into tests, and a doctest-style runner.

By the time of the decision the repository already had a version of extraction. `VisioAutomation_2010/VSamples.Docs` holds every C# block from the .NET docs, tagged with `// <sample id=... doc=... line=...>` markers, and `Check-DocsSamples.ps1 -DocsPath <docs checkout>` confirms each docs block matches its sample. The build compiles every sample. The last run covered 174 blocks, 171 samples and 3 exempt (API outlines that cannot compile as written), with no warnings.

A probe on [#250](https://github.com/saveenr/VisioAutomation/issues/250) showed that .NET 10 file-based apps could also run each sample against live Visio.

## Decision

Compile-and-compare is sufficient. `VSamples.Docs` plus `Check-DocsSamples.ps1` is the mechanism. Samples are not run automatically. The single-file `.cs` run-each-sample mechanism stays parked as a future experiment ([#250](https://github.com/saveenr/VisioAutomation/issues/250)).

## Why

- It catches the failure that actually happened: renames and removals (the 5.0.0 draw-method renames, `.Value` to `.Formula`).
- It already exists and costs nothing more to keep.
- Running samples needs live Visio, so it cannot run in CI, and it adds two header variants per sample (`#:package` for released docs, `#:project` for `next`). No sample is known to have compiled and still failed.

## Consequences

- The check is only as reliable as the habit of running it. CI has no docs checkout, so run the check script and build `VSamples.Docs` before each docs sweep at release.
- "Does it work" is not covered. The existing `VTest` and `VTest.Scripting` suites exercise most of the same APIs against real Visio; spot-check a few key samples by hand at release.
- The three exempt blocks are never checked.
- The PowerShell docs have no equivalent check. That gap is separate from this decision.

## Cross-references

- [`VisioAutomation_2010/VSamples.Docs/README.md`](../../VisioAutomation_2010/VSamples.Docs/README.md)
- [`tests-need-visio.md`](tests-need-visio.md)
- [#157](https://github.com/saveenr/VisioAutomation/issues/157) (docs location), [#250](https://github.com/saveenr/VisioAutomation/issues/250) (single-file run probe)

## Reconsider when

- A sample that compiles but fails when run reaches a released docs page.
- The docs move into this repository ([#157](https://github.com/saveenr/VisioAutomation/issues/157)), which would make an automatic check cheap enough to run in CI.
