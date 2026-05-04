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

**Phase 1 done. Visio PowerShell module 4.6.1 shipped to PSGallery on 2026-05-03** (tag `VisioPS_4.6.1` on `2026_Refresh`).

This release effectively pulled forward what had been planned for Phase 2: it bundles all the Phase 1 cleanup work plus four cmdlet bug fixes (`Lock-VisioShape` / `Unlock-VisioShape` switches now actually bind, `Export-VisioShape` no longer trips on its inverted file-existence check, `New-VisioShape` polyline / Bezier minimum-point validation actually throws). NuGet was deliberately not bumped — all four fixes are in `VisioPowerShell/Commands/`, not in the underlying library — so NuGet stays at `2.6.0` and the version-divergence policy decision is still deferred.

The first publish run surfaced several PSGallery / PS 5.1 gotchas (TLS 1.2 default, PowerShellGet 1.x silent-error bug, PS 5.1 vs 7 user-module path divergence, .ps1 file encoding). All are documented in [VisioPowerShellDocs/developer-info/publishing-to-powershell-gallery.md](https://saveenr.gitbook.io/visiopowershell/developer-info/publishing-to-powershell-gallery) and worked around by [`Publish-VisioPSToGallery.ps1`](VisioAutomation_2010/VisioPowerShell/Publish-VisioPSToGallery.ps1).

**Doc audit:** complete and pushed on both gitbook repos. See [`docs/AUDIT_PROGRESS.md`](docs/AUDIT_PROGRESS.md) for the by-section log; the remaining open items there are code-level findings (none Phase 1 blockers) and a list of out-of-audit-scope stub pages. The audit tracker can be deleted once `2026_Refresh` merges to `master`.

**Branch state:** the `2026_Refresh` branch is now pushed to origin and tagged. `master` has not yet been fast-forwarded to it — that's a pending decision (see Followups).

**Followups, recorded for the next session:**

1. **Set up GitHub Actions release CI** — automate the manual publish that was just done by hand. The current `Publish-VisioPSToGallery.ps1` carries the lessons (TLS, verification, etc.) and is the natural reference for the workflow. Tracked in [`docs/FUTURES.md`](docs/FUTURES.md) under *"Automate releases via GitHub CI"*.
2. **Merge `2026_Refresh` → `master`.** The release-tagged commit lives on the feature branch; convention is to fast-forward master after a release ships.
3. **Resume the open Phase-2-deferred items** when ready: version-number policy (NuGet `2.6.0` vs PS `4.6.1`), leftover-Visio-process flakiness investigation. Both still in [`docs/FUTURES.md`](docs/FUTURES.md).

## Other docs in this repo

- [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md) — projects, dependencies, central concepts
- [docs/BUILDING.md](docs/BUILDING.md) — full build / test / install reference (incl. dev-pack install commands)
- [docs/GLOSSARY.md](docs/GLOSSARY.md) — Visio and codebase terminology
- [docs/FUTURES.md](docs/FUTURES.md) — staged backlog of cleanup/modernization work
- [docs/AUDIT_PROGRESS.md](docs/AUDIT_PROGRESS.md) — in-progress doc-audit tracker (delete after Section C completes and merges)
- [docs/OVERVIEW.md](docs/OVERVIEW.md) — entry point, links to the above
