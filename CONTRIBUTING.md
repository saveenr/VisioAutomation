# Contributing to VisioAutomation

Thanks for your interest in contributing. This is a small, focused project — please read this short guide before opening a pull request.

## Active branch

Development is on **`master`**. Target it for pull requests, and consult [`docs/FUTURES.md`](docs/FUTURES.md) to see what's in scope for the current phase of the [2026 refresh](docs/FUTURES.md). (Phase 1 merged into `master` on 2026-05-03; Phase 2 / 3 work is upcoming.)

## Setup

Build prerequisites and exact commands: [`docs/BUILDING.md`](docs/BUILDING.md).

In short:
- Microsoft Visio installed locally (required to build *and* run the tests).
- Visual Studio 2022 (VS 2026 is not yet supported — Phase 3 work).
- A regular `git clone` and a build via the IDE or the documented `MSBuild.exe` invocation.

## Running the tests

All tests exercise real Visio COM calls. There is no mock layer (intentional — see [`docs/FUTURES.md`](docs/FUTURES.md) for context). You cannot run the tests on a machine without Visio installed.

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

## What's in scope right now

The 2026 refresh is staged in three phases (see [`docs/FUTURES.md`](docs/FUTURES.md)).

**Phase 1 *(current)*** accepts code and docs improvements with no new features, no TFM bumps, no IDE upgrades, no csproj-format changes, and no breaking API changes. PRs that violate those guardrails will be deferred to Phase 3.

**Phase 2** is the final release of the old-shape NuGet and PowerShell module.

**Phase 3** is the modernization (VS 2026, modern C#, possibly modern .NET, automated releases) and accepts changes that Phase 1 explicitly defers.

When in doubt about scope, open an issue first to discuss.

## Where to ask questions

Open a GitHub issue. There's no chat or mailing list.
