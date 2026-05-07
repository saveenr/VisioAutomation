# Futures — 2026 Refresh Backlog (index)

A running list of cleanup, modernization, and improvement items for the VisioAutomation solution. Items are grouped by theme into separate files (linked below). Each entry includes a one-line **What**, a **Why** (cost of leaving it), and a rough **Effort** (S / M / L). This is a *backlog* — items are not committed to or scheduled until pulled out into actual work.

> **Forward-looking only.** Done items live in [`COMPLETED.md`](COMPLETED.md). The phase-level "what shipped" headlines live in [`ROADMAP.md`](ROADMAP.md). When an item lands, the body moves to `COMPLETED.md` and is removed from its `futures/*.md` file. See [`CONTRIBUTING.md`](../CONTRIBUTING.md) for the convention.

---

## Where to find things

- **[`ROADMAP.md`](ROADMAP.md)** — staged-plan overview (Phase 1 / 2 / 3 status, what shipped per phase, what's still pending). Read this first for orientation.
- **[`futures/build-and-code.md`](futures/build-and-code.md)** — Build & tooling, Code & architecture. Items: *Consolidate target frameworks*, *Run tests in CI*, *Move development to Visual Studio 2026*, *Consider migrating off Visio 2010 PIA*, *Move to C# 14 / .NET 10*, *Make `CustomPropertyCells` values not require manual `EncodeValues()`*, *Borrow ideas from VisioBot3000 for VisioPS ergonomics*, *Borrow ideas from PSVA for VisioPS bulk-operation cmdlets*, *Evaluate NetOffice / NetOfficeFw as a replacement for the Visio PIA*, *Move `LinqExtensions` out of `Internal/`*.
- **[`futures/tests.md`](futures/tests.md)** — Test-related items. Items: *Tests require a live Visio* (design decision), *Test coverage gaps*, *Evaluate modern testing-stack options*.
- **[`futures/releases.md`](futures/releases.md)** — Release process and version policy. Items: *Reconcile version numbers across artifacts*, *Switch module-release builds from Debug to Release*, *Address `Visio.psd1` deprecation warnings on PSGallery publish*, *Automate releases via GitHub CI*.
- **[`futures/docs.md`](futures/docs.md)** — Documentation items, in-repo and user-facing gitbook. Items: *Decide where docs live long-term*, *Restructure the user-docs repos*, *Expand .NET-side doc coverage — Tier 3*, *Decide whether to document `VisioScripting` as a public API*, *Keep CHANGELOGs current*, plus five smaller gitbook items (including *Extend gitbook custom-properties page with Visio behavior matrix*).
- **[`futures/identity.md`](futures/identity.md)** — Dev team identity. Items: *Transition dev team identity from "Saveen" to "SevenPens"* (9 axes total; axes 1, 2, 3, 4, 6, 7, 8 done as of 2026-05-07. Axis 5 (hosting URLs) tracked in [#146](https://github.com/saveenr/VisioAutomation/issues/146) + [#147](https://github.com/saveenr/VisioAutomation/issues/147), scheduled `2026-12` milestone. Axis 9 (retire unused VisioAutomation legacy account) tracked in [#148](https://github.com/saveenr/VisioAutomation/issues/148), unscheduled).
