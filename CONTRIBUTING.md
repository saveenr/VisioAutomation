# Contributing to VisioAutomation

Thanks for your interest in contributing. This is a small, focused project — please read this short guide before opening a pull request.

## Active branch

Development is on **`master`**. Target it for pull requests, and consult [`docs/ROADMAP.md`](docs/ROADMAP.md) to see what's in scope for the current phase of the [2026 refresh](docs/ROADMAP.md). (Phase 1 merged into `master` on 2026-05-03; Phase 2 / 3 work is upcoming.)

## Setup

Build prerequisites and exact commands: [`docs/BUILDING.md`](docs/BUILDING.md).

In short:
- Microsoft Visio installed locally (required to build *and* run the tests).
- Visual Studio 2022 (VS 2026 is not yet supported — Phase 3 work).
- A regular `git clone` and a build via the IDE or the documented `MSBuild.exe` invocation.

## Running the tests

All tests exercise real Visio COM calls. There is no mock layer (intentional — see [`docs/decisions/tests-need-visio.md`](docs/decisions/tests-need-visio.md) for the full decision record). You cannot run the tests on a machine without Visio installed.

## Code style

- The codebase predates many modern C# conventions. **Don't reformat existing code** in a PR that's about something else — keep the diff focused on the actual change.
- New code should be reasonable C# 8 (the language version the projects compile with).
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

The release workflows ([`release-nuget.yml`](.github/workflows/release-nuget.yml), [`release-psmodule.yml`](.github/workflows/release-psmodule.yml)) source each release's GitHub Release notes from the matching CHANGELOG's `[Unreleased]` section. The workflow refuses to run if `[Unreleased]` is empty or still contains the placeholder `_No consumer-visible changes yet._` — populate it before triggering a release.

After a release lands, the maintainer converts the `[Unreleased]` section into a versioned section in the same CHANGELOG (per the Keep a Changelog convention):

```markdown
## [Unreleased]

_No consumer-visible changes yet._

## [2.6.1] - 2026-06-15

### Fixed
- ...
```

This is a manual post-release step — the workflows don't auto-edit the CHANGELOG.

## Backlog hygiene

The forward-looking backlog is split into topic files under [`docs/futures/`](docs/futures/) (indexed by [`docs/FUTURES.md`](docs/FUTURES.md)): `build-and-code.md`, `tests.md`, `releases.md`, `docs.md`. The phase-level "what shipped" headlines live in [`docs/ROADMAP.md`](docs/ROADMAP.md). Completed items' full detail lives in [`docs/COMPLETED.md`](docs/COMPLETED.md), grouped by phase. The split exists so each topic file stays scannable as a "what's left in this area" view while institutional memory (what was tried, why decisions were made, commit hashes) is preserved separately.

**When you finish a backlog item:**

1. Move the entry's body — the `**Resolution:**` paragraph and any sub-bullets — to the right phase + category section of `docs/COMPLETED.md`, verbatim.
2. Delete the body from the appropriate `docs/futures/*.md` file.
3. If the entry had a tail (a "Still to do" or "Deferred to Phase N" note pointing to follow-up work), extract that tail as a new active item in the appropriate `docs/futures/*.md` so the follow-up isn't lost.
4. Add a one-line bullet to the relevant phase's "items completed" checklist in `docs/ROADMAP.md` summarizing what shipped.

The headline phase summary in `ROADMAP.md` (e.g. "Phase 1 items completed:") and the body in `COMPLETED.md` play different roles: the headline is the project arc you scan to see "what got done"; `COMPLETED.md` is where you go when you actually need the detail.

## What's in scope right now

The 2026 refresh is staged in three phases (see [`docs/ROADMAP.md`](docs/ROADMAP.md)).

**Phase 1 *(current)*** accepts code and docs improvements with no new features, no TFM bumps, no IDE upgrades, no csproj-format changes, and no breaking API changes. PRs that violate those guardrails will be deferred to Phase 3.

**Phase 2** is the final release of the old-shape NuGet and PowerShell module.

**Phase 3** is the modernization (VS 2026, modern C#, possibly modern .NET, automated releases) and accepts changes that Phase 1 explicitly defers.

When in doubt about scope, open an issue first to discuss.

## Where to ask questions

Open a GitHub issue. There's no chat or mailing list.
