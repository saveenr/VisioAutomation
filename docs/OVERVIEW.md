# VisioAutomation — Documentation

**VisioAutomation** is a set of .NET libraries for automating a running instance of Microsoft Visio. It wraps the Visio COM API in ergonomic .NET types, adds a higher-level scripting facade, and ships a PowerShell module on top.

This `docs/` folder contains the **internal / developer-facing** documentation for the codebase itself. The user-facing usage docs live separately:

- VisioAutomation user guide — https://saveenr.gitbook.io/visioautomation/
- Visio PowerShell user guide — https://saveenr.gitbook.io/visiopowershell/
- Source for those gitbook docs — https://github.com/saveenr/VisioAutomation_GitBook_Docs

## Documents in this folder

- **[ARCHITECTURE.md](ARCHITECTURE.md)** — projects in the solution, what each is responsible for, how they depend on one another, and the central concepts (ShapeSheet addressing, batch I/O, the DOM, the scripting Client).
- **[BUILDING.md](BUILDING.md)** — prerequisites, how to build, how to run the tests, how to load the PowerShell module locally, and known rough edges worth cleaning up.
- **[TESTING.md](TESTING.md)** — design and conventions of the test suite: shared `Framework.VTest` base class, the per-testhost Visio singleton, `[AssemblyCleanup]` orphan-prevention, and the MSTEST0030 quality gate. Per-project READMEs live next to each test csproj.
- **[GLOSSARY.md](GLOSSARY.md)** — Visio-specific terms (ShapeSheet, SRC/SIDSRC, master, stencil, …) and codebase-specific terms (`VisioObjectTarget`, `Target*`, cell-group types, …).
- **[ROADMAP.md](ROADMAP.md)** — staged plan for the 2026 refresh (Phase 1 / 2 / 3 status, what shipped per phase, what's still pending). Read first for orientation.
- **[FUTURES.md](FUTURES.md)** — index of the backlog. Items are split by topic into [`futures/build-and-code.md`](futures/build-and-code.md), [`futures/tests.md`](futures/tests.md), [`futures/releases.md`](futures/releases.md), [`futures/docs.md`](futures/docs.md).

## Release history

Release notes for each shipped artifact live next to the artifact itself, in [Keep a Changelog](https://keepachangelog.com/en/1.1.0/) format:

- **[NuGet/CHANGELOG.md](../NuGet/CHANGELOG.md)** — `VisioAutomation2010` NuGet package
- **[VisioAutomation_2010/VisioPowerShell/CHANGELOG.md](../VisioAutomation_2010/VisioPowerShell/CHANGELOG.md)** — `Visio` PowerShell module
