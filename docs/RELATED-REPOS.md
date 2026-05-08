# Related repositories

A portfolio map of the Visio-related GitHub repositories under [`saveenr`](https://github.com/saveenr). The goal is for visitors to be able to tell at a glance which repos are active, which are paused or archive-only, and which are superseded by newer projects.

This index is the **Phase H1 deliverable** from [issue #152](https://github.com/saveenr/VisioAutomation/issues/152). Status calls were drafted from a metadata audit on 2026-05-07 and confirmed by the owner on 2026-05-08; per-repo `README.md` updates landed on the same day. Phase H2 (deeper retire / merge / maintain decisions per repo) is opportunistic and tracked separately; the Visio-PIAs source-consolidation question ([#180](https://github.com/saveenr/VisioAutomation/issues/180)) is the first bite of it.

## Status legend

- **active** &mdash; under regular development; the canonical place to file issues.
- **active reference** &mdash; load-bearing input to another active repo, but not itself receiving feature work.
- **paused** &mdash; feature-complete or last-shipped state is usable; not abandoned, but no near-term plans for active development.
- **superseded by X** &mdash; replaced by another repo that should be used instead. Source code retained for history.
- **abandoned** &mdash; no path forward planned. Distinct from "paused" in that even bug fixes are unlikely.

## Inventory

| Repo | Status | Stars | Last push | One-line description |
| :--- | :--- | ---: | :--- | :--- |
| [VisioAutomation](https://github.com/saveenr/VisioAutomation) | **active** | &mdash; | continuous | Primary repo. Ships the [`VisioAutomation2010`](https://www.nuget.org/packages/VisioAutomation2010/) NuGet and the [`Visio`](https://www.powershellgallery.com/packages/Visio) PowerShell module. |
| [Visio-PIAs](https://github.com/saveenr/Visio-PIAs) | **active reference** | 1 | 2017-02-03 | Source of the [`Visio2010.PrimaryInteropAssembly`](https://github.com/saveenr/VisioAutomation/blob/master/VisioAutomation_2010/Directory.Packages.props) NuGet package referenced from VisioAutomation. Not actively developed but load-bearing for VisioAutomation builds. Source-consolidation question tracked in [#180](https://github.com/saveenr/VisioAutomation/issues/180). |
| [VisioAutomation.VDX](https://github.com/saveenr/VisioAutomation.VDX) | **paused** | 13 | 2017-02-22 | Library to create simple Visio VDX files **without** Microsoft Visio installed. Distributed as the [`VisioAutomation.VDX`](https://www.nuget.org/packages/VisioAutomation.VDX/) NuGet package. Highest-starred sibling. CY27 hygiene pass (tests + docs refresh) tracked in [VisioAutomation.VDX#5](https://github.com/saveenr/VisioAutomation.VDX/issues/5). |
| [Visio-Power-Tools](https://github.com/saveenr/Visio-Power-Tools) | **paused** | 8 | 2017-01-19 | A set of Visio tools (`VisioPowerTools2010/`). Targets Visio 2010. CY27 revival (CI + VS 2022 + bump VA dependency) tracked in [Visio-Power-Tools#2](https://github.com/saveenr/Visio-Power-Tools/issues/2). |
| [Visio-Export-Pages-To-Docs](https://github.com/saveenr/Visio-Export-Pages-To-Docs) | **paused** | 12 | 2016-05-30 | Command-line tool that splits a multi-page Visio document into one document per page. Best-documented README of the paused-tool group. Requires Visio 2007 or above. CY27 docs cleanup tracked in [Visio-Export-Pages-To-Docs#5](https://github.com/saveenr/Visio-Export-Pages-To-Docs/issues/5). |
| [Visio-Code-Samples](https://github.com/saveenr/Visio-Code-Samples) | **paused, samples archive** | 3 | 2015-05-23 | Visio automation samples in CPython, .NET, IronPython, PowerShell, and VBA. CY27 VS 2022 compile pass tracked in [Visio-Code-Samples#1](https://github.com/saveenr/Visio-Code-Samples/issues/1). |
| [visio-reference](https://github.com/saveenr/visio-reference) | **paused, reference data** | 2 | 2017-01-23 | Plain-text dumps of stencils, masters, and templates that ship with Visio 2007 / 2010 / 2013. Useful as lookup data; no code. Newer-Visio-version extension question tracked in [visio-reference#1](https://github.com/saveenr/visio-reference/issues/1). |
| [VisioAutomation2007](https://github.com/saveenr/VisioAutomation2007) | **superseded by VisioAutomation** | 1 | 2017-08-17 | Older VisioAutomation source code targeting Visio 2007. The README itself describes the repo as an archive. |
| [Visio-Font-Compare](https://github.com/saveenr/Visio-Font-Compare) | **paused** | 1 | 2015-06-02 | WinForms tool to compare Visio fonts. Oldest sibling; initially classified as abandoned in the audit but reclassified to paused since a CY27 refresh is planned. CY27 refresh (VS 2022 + bump VA dependency) tracked in [Visio-Font-Compare#1](https://github.com/saveenr/Visio-Font-Compare/issues/1). |

The 9th sibling, `visio-templates`, was confirmed empty during the audit and was deleted by the owner on 2026-05-08.

## Caveats

- **License coverage is uneven.** Of the 8 siblings, only `VisioAutomation2007` and `VisioAutomation.VDX` have a `LICENSE` file in the repo root. The others have no GitHub-detected license. Phase H1 does not change that (it's a docs-only pass), but consumers should ask before depending on the unlicensed siblings.
- **Out of scope:** the gitbook docs repos for the .NET library and the PowerShell module ([`VisioAutomation_GitBook_Docs`](https://github.com/saveenr/VisioAutomation_GitBook_Docs), [`VisioPowerShellDocs`](https://github.com/saveenr/VisioPowerShellDocs)) are also under `saveenr`, but they're tooling for VisioAutomation rather than standalone projects, so they're not included in the table above. The issue body's repo list explicitly scoped them out.
- **Other `saveenr/visio-*` or related repos** that aren't listed here weren't named in [issue #152](https://github.com/saveenr/VisioAutomation/issues/152). If new ones turn up they can be added as Phase H1 follow-up rows.

## Status

Phase H1 acceptance criteria both met as of 2026-05-08:

1. ~~Centralized portfolio index.~~ **Done.** This file.
2. ~~Per-repo `README.md` updates on each sibling repo: `Status:` line, one-line description, successor pointer where applicable.~~ **Done** (8 commits, one per sibling repo, 2026-05-08).

Bonus cleanup: ~~delete the empty `visio-templates` repo.~~ **Done** by the owner on 2026-05-08.

CY27 work for paused repos is captured in the per-repo issues linked in the inventory above; the parent tracker is [#152](https://github.com/saveenr/VisioAutomation/issues/152). Phase H2 (retire / merge / maintain decisions per row) is opportunistic; gets filed as its own issue once the picture is firm. The Visio-PIAs source-consolidation question ([#180](https://github.com/saveenr/VisioAutomation/issues/180)) is the first bite.
