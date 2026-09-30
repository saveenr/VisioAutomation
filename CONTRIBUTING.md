# Contributing to VisioAutomation

Thanks for your interest in contributing. This is a small, focused project — please read this short guide before opening a pull request.

## Active branch

Development is on **`master`**. Target it for pull requests. Consult [`docs/MILESTONES.md`](docs/MILESTONES.md) for planned work and [`docs/HANDOVER.md`](docs/HANDOVER.md) for maintainer readiness.

## Setup

Build prerequisites and exact commands: [`docs/BUILDING.md`](docs/BUILDING.md).

In short:
- Microsoft Visio installed locally to run integration tests; compilation does not require Visio.
- Visual Studio 2026 and the .NET 10 SDK, the shared build-tool baseline with VisioAutomation.VDX. Runtime targets remain unchanged.
- A regular `git clone` and a build via the IDE or the documented `MSBuild.exe` invocation.

## Running the tests

The full suite exercises real Visio COM calls and needs Visio installed. Some metadata and pure-data tests can run without it. There is no mock Visio layer (intentional; see [`docs/decisions/tests-need-visio.md`](docs/decisions/tests-need-visio.md)). Run all four test assemblies using [`docs/BUILDING.md`](docs/BUILDING.md), and report failures and skips.

## Code style

- The codebase predates many modern C# conventions. **Don't reformat existing code** in a PR that's about something else — keep the diff focused on the actual change.
- Follow the language settings in `VisioAutomation_2010/Directory.Build.props` and preserve the existing .NET Framework targets.
- Don't add new files unless they're required by the change.
- Default to no comments. Only add one when the *why* is non-obvious. Identifier names should carry the *what*.

## Commit messages

- Subject line: concise, imperative, ≤ ~70 chars (`Fix X`, `Add Y`, `Update Z`).
- Body: explain the *why*, not the *what* — the diff already shows the *what*.
- Keep one logical change per commit. If you find yourself writing "and also …" in the subject, split the commit.

## Changelogs

The project ships two artifacts that consumers depend on:

- The [`VisioAutomation2010`](NuGet/CHANGELOG.md) NuGet package
- The [`Visio`](VisioAutomation_2010/VisioPowerShell/CHANGELOG.md) PowerShell module

When your change is **consumer-visible** (public API, behavior, supported runtime, dependencies), add an entry to the matching `[Unreleased]` section of the corresponding `CHANGELOG.md` in the **same commit**, following the [Keep a Changelog 1.1.0](https://keepachangelog.com/en/1.1.0/) format already in use.

Pure internal / build / docs changes don't need changelog entries.

### Release flow

The release workflows ([`release-nuget.yml`](.github/workflows/release-nuget.yml), [`release-psmodule.yml`](.github/workflows/release-psmodule.yml)) read notes from the matching CHANGELOG's **versioned** `[<version>]` section. They fail if that section is missing or empty.

In the version-bump commit, **before triggering a release**, move the `[Unreleased]` entries into a versioned section and create a fresh `[Unreleased]` section:

```markdown
## [Unreleased]

_No consumer-visible changes yet._

## [2.6.1] - 2026-06-15

### Fixed
- ...
```

This is a manual release-preparation step; workflows do not edit the CHANGELOG. Build and test Release, then use the release workflow followed by its publish workflow. See [`docs/BUILDING.md`](docs/BUILDING.md).

## Backlog hygiene

The forward-looking backlog is split into topic files under [`docs/futures/`](docs/futures/) (indexed by [`docs/FUTURES.md`](docs/FUTURES.md)): `build-and-code.md`, `tests.md`, `releases.md`, `docs.md`. The phase-level "what shipped" headlines live in [`docs/ROADMAP.md`](docs/ROADMAP.md). Completed items' full detail lives in [`docs/COMPLETED.md`](docs/COMPLETED.md), grouped by phase. The split exists so each topic file stays scannable as a "what's left in this area" view while institutional memory (what was tried, why decisions were made, commit hashes) is preserved separately.

**When you finish a backlog item:**

1. Move the entry's body — the `**Resolution:**` paragraph and any sub-bullets — to the right phase + category section of `docs/COMPLETED.md`, verbatim.
2. Delete the body from the appropriate `docs/futures/*.md` file.
3. If the entry had a tail (a "Still to do" or "Deferred to Phase N" note pointing to follow-up work), extract that tail as a new active item in the appropriate `docs/futures/*.md` so the follow-up isn't lost.
4. Add a one-line bullet to the relevant phase's "items completed" checklist in `docs/ROADMAP.md` summarizing what shipped.

The headline phase summary in `ROADMAP.md` (e.g. "Phase 1 items completed:") and the body in `COMPLETED.md` play different roles: the headline is the project arc you scan to see "what got done"; `COMPLETED.md` is where you go when you actually need the detail.

## What's in scope right now

Handover work prioritizes reproducible builds and releases, reliable tests, and accurate documentation. SDK-style projects, central package management, and release automation are already implemented. Historical phase summaries live in [`docs/ROADMAP.md`](docs/ROADMAP.md); use [`docs/MILESTONES.md`](docs/MILESTONES.md) for the forward plan.

Runtime upgrades and public API changes need compatibility review. The [`VisioScripting` API contract](docs/decisions/visioscripting-public-api.md) describes the stable facade. Keep unrelated refactors out of readiness fixes.

## Where to ask questions

Open a GitHub issue. There's no chat or mailing list.
