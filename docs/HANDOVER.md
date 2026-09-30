# Maintainer handover

Readiness assessment dated **2026-09-29**, against source commit `9653b207` plus the accompanying handover changes. This records local verification, not a new public release or a completed account transfer.

## Verified baseline

The full solution builds with VS 2022 MSBuild. Release integration tests pass against installed Visio 16.0.20326.20158 using the x64 VSTest runner. The module loads in Windows PowerShell 5.1.26100.9444 with all 64 cmdlets exported.

| Check | Result | Local evidence |
|---|---|---|
| Original Debug suite | 235 passed, 1 skipped | `TestResults/handover-baseline.trx` |
| Release suite after fixes | 237 passed, 0 failed, 0 skipped | `TestResults/handover-release.trx` |
| Visio cleanup | No Visio processes before or after completed full runs | Process inventory |
| NuGet packaging | Package created; all six `lib/net452` DLLs match Release output by SHA-256; README present | `TestResults/VisioAutomation2010.3.0.0.nupkg` |
| Release module | Manifest valid; exports match the assembly; optimization enabled | Windows PowerShell 5.1 import and reflection check |
| PowerShell syntax | Ten tracked scripts/manifests parse in Windows PowerShell 5.1 | Parser check |
| CI subset | Three metadata/session tests pass in each of Debug and Release | `TestResults/ci-module-checks-*.trx` |
| Workflow syntax | Five YAML files and 29 embedded PowerShell blocks parse | Local parser checks; hosted execution remains unverified |
| Documentation links | Relative file targets checked across all three repositories | 246 Markdown files; external URLs and anchors excluded |

Test counts by project: `VTest` 107, `VTest.Models` 60, `VTest.Scripting` 43, `VTest.PowerShell` 27. Build output still contains obsolete-API warnings from tests using the compatibility `Value` aliases. No warning suppression was added.

`TestResults` is ignored by Git. Retain the TRX files with the handover or release evidence; regenerate them using [BUILDING.md](BUILDING.md). The local package above retains the checked-in version solely for verification and must not be mistaken for a new published release.

## Fixes in this readiness pass

- NuGet releases and raw-DLL archives now use Release binaries, matching the PowerShell release flow.
- The PowerShell test harness imports the checkout's module before opening its runspace. A regression test disables autoloading and checks the assembly path, preventing accidental testing of an installed module.
- Scripted and directly invoked test cmdlets execute on the same thread so shared Visio COM objects retain their expected identity.
- The previously ignored export-overwrite test now verifies a valid PNG. Negative tests assert the expected error text instead of accepting any exception; the old export test could pass on a COM identity error.
- CI builds Debug and Release and runs the three module metadata/session checks without Visio. Full integration testing remains a local prerequisite.
- Build, contributor, test, and public publishing guidance has been aligned with the implemented workflows.
- Public XML-loader examples use the current facade, with explicit instructions for users of the published 3.0.0 package.

## Repository and service map

| Component | Source or configuration | Incoming maintainer responsibility |
|---|---|---|
| Libraries, module, tests, samples, CI | `VisioAutomation`, development branch `master` | Repository administration and workflow access |
| .NET user guide | Sibling `VisioAutomation_GitBook_Docs`, branch `main` | Repository and GitBook publishing access |
| PowerShell user guide | Sibling `VisioPowerShellDocs`, branch `visiops_v4_docs` | Repository and GitBook publishing access; preserve the correct docs branch |
| NuGet package | `NuGet/VisioAutomation2010.nuspec`; `release-nuget.yml` then `publish-nuget.yml` | Package ownership and `NUGET_API_KEY` |
| PowerShell Gallery module | `VisioAutomation_2010/VisioPowerShell/Visio.psd1`; `release-psmodule.yml` then `publish-psmodule.yml` | Module ownership and `PSGALLERY_API_KEY` |
| Visio PIA dependency | `Visio2010.PrimaryInteropAssembly` in `Directory.Packages.props` | Know its provenance and package availability; see [RELATED-REPOS.md](RELATED-REPOS.md) |

This pass did not inspect or change external ownership, secrets, GitBook settings, branch protections, published packages, or GitHub issues. Existing `saveenr` and `SevenPens` references describe historical hosting/publishing identities, not the identity of the incoming owner.

## Before the next release or transfer

1. **Resolve release compatibility.** The NuGet `[Unreleased]` changelog includes formerly public types and methods becoming internal. Historical plans call the next version `3.1.0`, but callers using those APIs can break. Decide whether to preserve compatibility or use an appropriate breaking-release version before publishing; do not infer safety merely because the internal tests pass. See [the API contract](decisions/visioscripting-public-api.md) and [NuGet changelog](../NuGet/CHANGELOG.md).
2. **Verify the new maintainer's access.** Confirm repository administration for all three repositories, GitBook access, and ownership of both public packages. Have the incoming publisher configure their own scoped keys in the two named repository secrets. The keys themselves do not belong in this document.
3. **Prepare and rehearse the release.** Bump versions, roll changelog entries into versioned sections, run the full Release suite, and exercise both release workflows with `dry_run`. Publish workflows validate an already-created GitHub Release artifact; their dry runs cannot prove the final upload credentials work. Hosted workflow execution has not been performed in this local pass.
4. **Verify documentation publication after any move.** Confirm GitBook sync branches and access, then update repository/package/documentation links for the chosen destination. Local relative-file checks do not verify external URLs, redirects, rendered GitBook pages, or all sample code.
5. **Agree on the support boundary.** This verification covers one modern Visio installation and x64 execution. It does not establish a Visio 2010-through-current, 32-bit, or PowerShell 7 compatibility matrix. Runtime upgrades and broader cmdlet coverage remain separate work in [MILESTONES.md](MILESTONES.md) and [the coverage audit](futures/test-coverage-gaps.md).

## First maintenance session

Read [ARCHITECTURE.md](ARCHITECTURE.md), build and run all four suites with [BUILDING.md](BUILDING.md), then import the local Release manifest in a fresh Windows PowerShell 5.1 session. Use [TESTING.md](TESTING.md) when adding regression tests and [CONTRIBUTING.md](../CONTRIBUTING.md) for change and changelog conventions.

Older roadmap and session notes preserve useful history, but their dates, counts, completed phases, and prospective version numbers are not current verification. Prefer the source configuration, reproducible commands, and dated results above.
