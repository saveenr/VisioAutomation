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

## Build commands (verified working)

```bash
MSBUILD="/c/Program Files/Microsoft Visual Studio/2022/Community/MSBuild/Current/Bin/MSBuild.exe"
"$MSBUILD" VisioAutomation_2010/VisioAutomation2010.sln -t:Restore -p:RestorePackagesConfig=true
"$MSBUILD" VisioAutomation_2010/VisioAutomation2010.sln -p:Configuration=Debug -m
```

- **Do not** use `dotnet build` — projects are legacy csproj + packages.config.
- **Do not** use VS 2026's MSBuild (under `Program Files\Microsoft Visual Studio\18\`) — its .NET Framework floor is 4.6.2 and the shipping libs are on 4.5. This is a Phase 3 unblocker.

Full reference (IDE flow, test invocation, the v4.5 reference-assemblies NuGet workaround): [docs/BUILDING.md](docs/BUILDING.md).

## Tests need a live Visio

All test projects exercise real Visio COM calls. There is no mock/fake layer (intentional). Tests cannot be run on a machine without Microsoft Visio installed — flag this rather than claiming a green run.

## Per-commit conventions

- **Changelogs:** when a change is consumer-visible (public API, behavior, supported runtime, dependencies) add an entry to the matching `[Unreleased]` section of [`NuGet/CHANGELOG.md`](NuGet/CHANGELOG.md) or [`VisioAutomation_2010/VisioPowerShell/CHANGELOG.md`](VisioAutomation_2010/VisioPowerShell/CHANGELOG.md) in the **same commit**. Pure internal / build / docs changes don't need entries.
- **PowerShell loader scripts** in `VisioAutomation_2010/VisioPowerShell/`: `Load*` does in-session imports; `Install*` does persistent installs to the user's PS modules folder; `*.ISE.ps1` is the ISE-launched variant. See the folder's [README.md](VisioAutomation_2010/VisioPowerShell/README.md).

## Tooling notes

- **Shell:** Windows host. Both Bash and PowerShell are available. Use Bash for git and Unix-style tooling; use PowerShell for `.ps1` parse checks (`[System.Management.Automation.PSParser]::Tokenize`) and Windows-specific operations.
- **`gh`** is not yet relied on by this repo's workflows. The user-facing gitbook docs live in a separate repo: [VisioAutomation_GitBook_Docs](https://github.com/saveenr/VisioAutomation_GitBook_Docs). Don't try to edit them through this repo.

## Other docs in this repo

- [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md) — projects, dependencies, central concepts
- [docs/BUILDING.md](docs/BUILDING.md) — full build / test / install reference
- [docs/GLOSSARY.md](docs/GLOSSARY.md) — Visio and codebase terminology
- [docs/FUTURES.md](docs/FUTURES.md) — staged backlog of cleanup/modernization work
- [docs/OVERVIEW.md](docs/OVERVIEW.md) — entry point, links to the above
