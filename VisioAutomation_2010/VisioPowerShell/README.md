# VisioPowerShell

The PowerShell module front-end for VisioAutomation. Builds to `VisioPS.dll`, distributed as the `Visio` module on the [PowerShell Gallery](https://www.powershellgallery.com/packages/Visio).

For the bigger picture — how this project relates to the rest of the solution, its dependencies, and the cmdlet → VisioScripting flow — see [`docs/ARCHITECTURE.md`](../../docs/ARCHITECTURE.md).

## Folder layout

- `Commands/` — cmdlet implementations grouped by noun (`VisioApplication/`, `VisioShape/`, `VisioPage/`, `VisioCustomProperty/`, …). Each `.cs` file holds one verb-noun cmdlet class (`Get-VisioShape`, `New-VisioShape`, `Set-VisioShapeCells`, …).
- `Models/` — small types used as cmdlet parameter and return shapes (`PageCells`, `ShapeCells`, `BaseCells`, …).
- `Internal/` — non-public helpers (`CellTuple`, `NameValueDictionary`, `NamedSrcDictionary`, `DataTableHelpers`).
- `Demo/` — interactive demo runner and sample data; see below.
- `Visio.psd1` — module manifest (declares cmdlet exports, type files, required assemblies).
- `Visio.Types.ps1xml` — display formatting for Visio COM objects in the pipeline.
- `VisioPsClientContext.cs` — bridges PowerShell's `Write*` methods into VisioScripting's `ClientContext` abstraction.

## Helper scripts

| Script | Purpose | When to use |
|---|---|---|
| [`LoadFromBinDebug.ps1`](LoadFromBinDebug.ps1) | Imports the freshly built `bin/Debug` module into the current session | Fastest dev iteration: build, then `. .\LoadFromBinDebug.ps1` |
| [`LoadFromBinDebug.ISE.ps1`](LoadFromBinDebug.ISE.ps1) | Launches PowerShell ISE running the script above | When you prefer ISE for testing |
| [`LoadFromGallery.ps1`](LoadFromGallery.ps1) | `Save-Module`s the published `Visio` module from PSGallery into a local `DownloadedModule/` (gitignored) and imports it | Verify that a published release loads cleanly in a clean session, separate from the local build |
| [`InstallForCurrentUser.ps1`](InstallForCurrentUser.ps1) | Robocopies `bin/Debug` into the user's `Documents/WindowsPowerShell/Modules/Visio/` folder | When you want any future PowerShell session to be able to `Import-Module Visio` from the local build |

### Naming convention

- `Load*` scripts perform an **in-session import** — transient, per-session, no files copied outside the repo.
- `Install*` scripts perform a **persistent install** — copies files into the user's PS modules folder.
- `*.ISE.ps1` indicates the PowerShell-ISE-launched variant of a script.

Keep new helper scripts consistent with this convention.

## Demo

The `Demo/` subfolder contains an interactive demo runner:

- `start-demo.ps1` — Joel "Jaykul" Bennett's classic `Start-Demo` script (third-party, public domain). Reads a sequence of commands from a text file and walks through them with single-keypress navigation.
- `demo.txt` — the command sequence the runner walks through.
- `directedgraph1.xml`, `directedgraph2.xml`, `orgchart1.xml` — sample inputs the demo loads.

Run from inside the `Demo/` folder, after the module is imported (e.g., via `LoadFromBinDebug.ps1`):

```powershell
.\start-demo.ps1
```

## See also

- [`CHANGELOG.md`](CHANGELOG.md) — release notes for the `Visio` PowerShell module
- [`docs/ARCHITECTURE.md`](../../docs/ARCHITECTURE.md) — solution-wide architecture and dependencies
- [`docs/BUILDING.md`](../../docs/BUILDING.md) — how to build, test, and install
- [`docs/GLOSSARY.md`](../../docs/GLOSSARY.md) — Visio and codebase terminology
- [`docs/ROADMAP.md`](../../docs/ROADMAP.md) — staged plan (Phase 1 / 2 / 3) for the 2026 refresh
- [`docs/FUTURES.md`](../../docs/FUTURES.md) — index of the topic-split backlog of planned cleanup and modernization work
