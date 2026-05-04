# CLAUDE.md

Project-specific guidance for Claude Code sessions in this repo. Loaded automatically.

## What this is

A .NET Framework library plus a PowerShell module that automate Microsoft Visio via COM interop. Full picture: [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md).

## The 2026 refresh — read this before making changes

Active branch: `2026_Refresh`. Work is staged in three phases per [docs/FUTURES.md](docs/FUTURES.md):

1. **Phase 1 — VS 2022 cleanup** *(currently in progress)*. Stay on the current TFMs (.NET Framework 4.5.2 for shipping libs, .NET Framework 4.7.2 for tests). **No new features.** No major TFM bumps, no IDE upgrades, no csproj-format changes — those wait for Phase 3.
2. **Phase 2 — Cut a final release** of the `VisioAutomation2010` NuGet and the `Visio` PowerShell module with refreshed docs.
3. **Phase 3 — Modernization.** Move to VS 2026 (which requires bumping TFMs to 4.7.2), modern C#, possibly modern .NET, automated releases.

When in doubt whether a change fits the current phase, check [docs/FUTURES.md](docs/FUTURES.md).

## Build prerequisites

The shipping libs target .NET Framework 4.5.2. Modern Windows install media don't include the v4.5.2 reference assemblies, so **the .NET Framework 4.5.2 Developer Pack must be installed** on every dev machine and CI runner. Microsoft does not publish a winget manifest for 4.5.x — use chocolatey (`choco install netfx-4.5.2-devpack -y`) or the Microsoft download page. Full instructions in [docs/BUILDING.md](docs/BUILDING.md).

If you see `MSB3644: The reference assemblies for .NETFramework,Version=v4.5.2 were not found` — that's the dev pack not being installed.

## Build commands (verified working)

```bash
MSBUILD="/c/Program Files/Microsoft Visual Studio/2022/Community/MSBuild/Current/Bin/MSBuild.exe"
"$MSBUILD" VisioAutomation_2010/VisioAutomation2010.sln -t:Restore -p:RestorePackagesConfig=true
"$MSBUILD" VisioAutomation_2010/VisioAutomation2010.sln -p:Configuration=Debug -m
```

- **Do not** use `dotnet build` — projects are legacy csproj + packages.config.
- **Do not** use VS 2026's MSBuild (under `Program Files\Microsoft Visual Studio\18\`) — its .NET Framework floor is 4.6.2 and the shipping libs are on 4.5.2. This is a Phase 3 unblocker.

Full reference (IDE flow, test invocation, dev-pack install commands): [docs/BUILDING.md](docs/BUILDING.md).

## Tests need a live Visio

All test projects exercise real Visio COM calls. There is no mock/fake layer (intentional). Tests cannot be run on a machine without Microsoft Visio installed — flag this rather than claiming a green run.

## Per-commit conventions

- **Changelogs:** when a change is consumer-visible (public API, behavior, supported runtime, dependencies) add an entry to the matching `[Unreleased]` section of [`NuGet/CHANGELOG.md`](NuGet/CHANGELOG.md) or [`VisioAutomation_2010/VisioPowerShell/CHANGELOG.md`](VisioAutomation_2010/VisioPowerShell/CHANGELOG.md) in the **same commit**. Pure internal / build / docs changes don't need entries.
- **PowerShell loader scripts** in `VisioAutomation_2010/VisioPowerShell/`: `Load*` does in-session imports; `Install*` does persistent installs to the user's PS modules folder; `*.ISE.ps1` is the ISE-launched variant. See the folder's [README.md](VisioAutomation_2010/VisioPowerShell/README.md).

## Tooling notes

- **Shell:** Windows host. Both Bash and PowerShell are available. Use Bash for git and Unix-style tooling; use PowerShell for `.ps1` parse checks (`[System.Management.Automation.PSParser]::Tokenize`) and Windows-specific operations.
- **GitHub access:** the `TheSevenPens` git identity has push access to all three repos in play (this repo, [`VisioAutomation_GitBook_Docs`](https://github.com/saveenr/VisioAutomation_GitBook_Docs), and [`VisioPowerShellDocs`](https://github.com/saveenr/VisioPowerShellDocs)). The user-facing docs live in those last two repos as siblings of this repo (cloned to `C:\Users\savee\Documents\GitHub\VisioAutomation_GitBook_Docs\` and `C:\Users\savee\Documents\GitHub\VisioPowerShellDocs\`). PS docs use a version-pinned branch (`visiops_v4_docs`), not master — see [reference_doc_repos.md memory](../../.claude/projects/C--Users-savee-Documents-GitHub-VisioAutomation/memory/reference_doc_repos.md).

## Current state (resume here)

**Phase 1 work is mostly done.** Everything from the original quick-wins list has landed: per-project READMEs, CONTRIBUTING.md, root readme rewrite, CLAUDE.md (this file), MSTest off the beta, PowerShell loader-script renames, per-artifact CHANGELOGs, dead-code cleanup in `Internal/`, build-only CI workflow, and TFM consolidation (now at .NET 4.5.2 for libs, .NET 4.7.2 for tests).

**Doc audit status** in [`docs/AUDIT_PROGRESS.md`](docs/AUDIT_PROGRESS.md):
- Section A (PS docs strict-accuracy fixes): **done, pushed**
- Section B (.NET docs strict-accuracy fixes): **done, pushed**
- Section C.7/C.8/C.9 (.NET docs content rewrites): **done, pushed**
- Section C, items 12–18 (PS docs *new* cmdlet pages — `New-VisioShape`, `Remove-VisioShape`, `New-/Set-VisioPageCells`, `New-/Get-VisioShapeCells`, the Control family section, `cmdlets/container.md` flesh-out, `cmdlets/other-cmdlets.md`): **done locally; 7 commits ahead of `origin/visiops_v4_docs`, not pushed**

The audit is now substantively complete. Remaining work flagged in `AUDIT_PROGRESS.md` (bare-headline stubs for `Copy-VisioShape`, `Lock-VisioShape`, etc.) is out of audit scope.

**Branch state:** the `2026_Refresh` branch on this repo is **local-only** — no upstream configured, ~30 commits not pushed. The `VisioPowerShellDocs` repo's `visiops_v4_docs` branch is 7 commits ahead of origin. The `VisioAutomation_GitBook_Docs` repo is pushed. Don't push `2026_Refresh` or the unpushed gitbook commits without explicit user confirmation.

**Phase 2 prerequisites that are deferred and need user discussion** (in [`docs/FUTURES.md`](docs/FUTURES.md)):
- Reconcile version numbers across artifacts (NuGet `2.6.0` vs PS module `4.6.0`)
- Investigate flakiness from leftover Visio processes

## Other docs in this repo

- [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md) — projects, dependencies, central concepts
- [docs/BUILDING.md](docs/BUILDING.md) — full build / test / install reference (incl. dev-pack install commands)
- [docs/GLOSSARY.md](docs/GLOSSARY.md) — Visio and codebase terminology
- [docs/FUTURES.md](docs/FUTURES.md) — staged backlog of cleanup/modernization work
- [docs/AUDIT_PROGRESS.md](docs/AUDIT_PROGRESS.md) — in-progress doc-audit tracker (delete after Section C completes and merges)
- [docs/OVERVIEW.md](docs/OVERVIEW.md) — entry point, links to the above
