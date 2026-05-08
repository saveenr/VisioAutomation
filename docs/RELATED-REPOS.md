# Related repositories

A portfolio map of the Visio-related GitHub repositories under [`saveenr`](https://github.com/saveenr). The goal is for visitors to be able to tell at a glance which repos are active, which are paused or archive-only, and which are superseded by newer projects.

This index is the **Phase H1 deliverable** from [issue #152](https://github.com/saveenr/VisioAutomation/issues/152). The status calls below were drafted from a metadata audit and are marked **(preliminary)** where they still need confirmation; the README on each sibling repo will be updated to match the final calls in a follow-up pass. Phase H2 (deeper retire / merge / maintain decisions per repo) is opportunistic and tracked separately.

## Status legend

- **active** — under regular development; the canonical place to file issues.
- **active reference** — load-bearing input to another active repo, but not itself receiving feature work.
- **paused** — feature-complete or last-shipped state is usable; not abandoned, but no near-term plans for active development.
- **superseded by X** — replaced by another repo that should be used instead. Source code retained for history.
- **empty / needs-decision** — repo exists but holds no content; future is undecided.
- **abandoned** — no path forward planned. Distinct from "paused" in that even bug fixes are unlikely.

## Inventory

| Repo | Status (preliminary unless noted) | Stars | Last push | One-line description |
| :--- | :--- | ---: | :--- | :--- |
| [VisioAutomation](https://github.com/saveenr/VisioAutomation) | **active** | &mdash; | continuous | Primary repo. Ships the [`VisioAutomation2010`](https://www.nuget.org/packages/VisioAutomation2010/) NuGet and the [`Visio`](https://www.powershellgallery.com/packages/Visio) PowerShell module. |
| [Visio-PIAs](https://github.com/saveenr/Visio-PIAs) | **active reference** *(preliminary)* | 1 | 2017-02-03 | Source of the [`Visio2010.PrimaryInteropAssembly`](https://github.com/saveenr/VisioAutomation/blob/master/VisioAutomation_2010/Directory.Packages.props) NuGet package referenced from VisioAutomation. Not actively developed but load-bearing for VisioAutomation builds. |
| [VisioAutomation.VDX](https://github.com/saveenr/VisioAutomation.VDX) | **paused** *(preliminary)* | 13 | 2017-02-22 | Library to create simple Visio VDX files **without** Microsoft Visio installed. Distributed as the [`VisioAutomation.VDX`](https://www.nuget.org/packages/VisioAutomation.VDX/) NuGet package. Highest-starred sibling. |
| [Visio-Power-Tools](https://github.com/saveenr/Visio-Power-Tools) | **paused** *(preliminary)* | 8 | 2017-01-19 | A set of Visio tools (`VisioPowerTools2010/`). Targeting Visio 2010. README is empty. |
| [Visio-Export-Pages-To-Docs](https://github.com/saveenr/Visio-Export-Pages-To-Docs) | **paused** *(preliminary)* | 12 | 2016-05-30 | Command-line tool that splits a multi-page Visio document into one document per page. Best-documented README of the paused-tool group. Requires Visio 2007 or above. |
| [Visio-Code-Samples](https://github.com/saveenr/Visio-Code-Samples) | **paused, samples archive** *(preliminary)* | 3 | 2015-05-23 | Visio automation samples in CPython, .NET, IronPython, PowerShell, and VBA. README is empty. |
| [visio-reference](https://github.com/saveenr/visio-reference) | **paused, reference data** *(preliminary)* | 2 | 2017-01-23 | Plain-text dumps of stencils, masters, and templates that ship with Visio 2007 / 2010 / 2013. Useful as lookup data; no code. |
| [VisioAutomation2007](https://github.com/saveenr/VisioAutomation2007) | **superseded by VisioAutomation** *(self-described, confirmed)* | 1 | 2017-08-17 | Older VisioAutomation source code targeting Visio 2007. The README itself describes the repo as an archive. |
| [Visio-Font-Compare](https://github.com/saveenr/Visio-Font-Compare) | **abandoned** *(preliminary, lowest-confidence)* | 1 | 2015-06-02 | WinForms tool to compare Visio fonts. Oldest sibling, never re-touched, README is just a header. |
| [visio-templates](https://github.com/saveenr/visio-templates) | **empty / needs-decision** *(confirmed empty)* | 1 | 2017-05-08 | Empty repo (no default branch, no contents). Likely a placeholder that was never populated; a candidate for either deletion or a clear "this never had content" README. |

## Caveats

- **Status calls are preliminary** for everything except VisioAutomation itself. The plan is for the user (the repo owner) to confirm or correct each status, after which each sibling repo's `README.md` will be updated with a matching `Status: ...` line and any successor-project pointer.
- **License coverage is uneven.** Of the 9 siblings, only `VisioAutomation2007` and `VisioAutomation.VDX` have a `LICENSE` file in the repo root. The others have no GitHub-detected license. Phase H1 does not change that (it's a docs-only pass), but consumers should ask before depending on the unlicensed siblings.
- **Out of scope:** the gitbook docs repos for the .NET library and the PowerShell module ([`VisioAutomation_GitBook_Docs`](https://github.com/saveenr/VisioAutomation_GitBook_Docs), [`VisioPowerShellDocs`](https://github.com/saveenr/VisioPowerShellDocs)) are also under `saveenr`, but they're tooling for VisioAutomation rather than standalone projects, so they're not included in the table above. The issue body's repo list explicitly scoped them out.
- **Other `saveenr/visio-*` or related repos** that aren't listed here weren't named in [issue #152](https://github.com/saveenr/VisioAutomation/issues/152). If new ones turn up they can be added as Phase H1 follow-up rows.

## Next steps

1. Owner confirms or corrects the status column above.
2. Per-repo `README.md` updates land on each sibling repo: a `Status: ...` line, the one-line description verbatim from this table, and a successor-project pointer where applicable.
3. Phase H2 (retire / merge / maintain decisions per row) gets filed as its own issue once the picture is firm.
