# Roadmap — 2026 Refresh

The 2026 refresh runs in three phases. Each phase has a one-line headline summary that stays here; full per-item detail moves to [`COMPLETED.md`](COMPLETED.md) when an item lands.

For the still-open backlog of work, see the topic-specific files under [`futures/`](futures/) — the [`FUTURES.md`](FUTURES.md) index lists them all.

---

## Phase 1 — VS 2022 cleanup *(done; merged to master 2026-05-03)*

Stayed on Visual Studio 2022 and the current TFMs (.NET Framework 4.5.2 for shipping libs, 4.7.2 for tests). Code + docs improvements only, no new features. The phase culminated in the **Visio PowerShell 4.6.1** release on 2026-05-03 (tag `VisioPS_4.6.1`).

Phase 1 items completed:
- ✅ *Revise user-facing documentation for accuracy* — full audit and rewrite of [VisioPowerShellDocs](https://saveenr.gitbook.io/visiopowershell) and the .NET-side gitbook docs. Standardized every cmdlet page on a Syntax + Parameters + Examples + See-also layout. Reader-facing summary at [`documentation-changes.md`](https://saveenr.gitbook.io/visiopowershell/documentation-changes).
- ✅ *Cmdlet bug fixes shipped in 4.6.1* — `Lock-VisioShape` / `Unlock-VisioShape` switches now actually bind; `Export-VisioShape` file-exists check no longer inverted; `New-VisioShape` polyline / Bezier minimum-point validation actually throws.
- ✅ *Manual release machinery* — [`Publish-VisioPSToGallery.ps1`](../VisioAutomation_2010/VisioPowerShell/Publish-VisioPSToGallery.ps1) wraps the staging / publish / tag / push flow with TLS 1.2 forcing, `-ErrorAction Stop`, and post-publish gallery verification. Documented in [VisioPowerShellDocs/developer-info/publishing-to-powershell-gallery.md](https://saveenr.gitbook.io/visiopowershell/developer-info/publishing-to-powershell-gallery).
- ✅ *Fix the misnamed PowerShell loader script* — rewrote it to actually `Save-Module` from the PS Gallery
- ✅ *Add a `CLAUDE.md` at the repo root* — added with staged-plan, build commands, conventions, doc pointers
- ✅ *Update MSTest off the beta* — upgraded `MSTest.TestFramework` and `MSTest.TestAdapter` from `2.0.0-beta2` to `4.2.2`; bumped `VTest` TFM 4.5 → 4.7.2 to satisfy MSTest 4.x's floor
- ✅ *Add a per-project `README.md` for the larger projects* — `VisioAutomation/`, `VisioAutomation.Models/`, `VisioScripting/`, `VisioPowerShell/` (already had one)
- ✅ *Add a `CONTRIBUTING.md`* — covers branch, setup pointer, tests-need-Visio rule, code style, commits, changelog discipline, per-phase scope
- ✅ *Expand the root `readme.md`* — rewrote with pitch, install table, C# + PowerShell quick-start, doc links, license
- ✅ *Audit `Internal/` for dead code* — deleted orphaned `TempHelper.cs` + removed dead `InternalsVisibleTo("TestVisioAutomation")` attribute; spawned a follow-up item for misc warts found during the audit
- ✅ *Misc cleanups discovered during the Internal/ audit* (mostly) — moved misplaced `InternalsVisibleTo` attributes to `AssemblyInfo.cs`, deleted two orphaned VTest files, removed auto-generated `.sln.metaproj` from version control. `LinqExtensions` visibility-vs-folder mismatch deferred to Phase 3 as a breaking-namespace-change risk.
- ✅ *Add CI* (build-only) — `.github/workflows/build.yml` builds the solution on push/PR for `master`, pinned to VS 2022 MSBuild, NuGet packages cached. Test runs in CI deferred to Phase 3 (needs self-hosted runner with Visio).

## Phase 2 — Cut the final release

Tag and publish a final release of VisioAutomation (NuGet) with the refreshed docs. The PowerShell-module half of this phase shipped early as **Visio PowerShell 4.6.1** on 2026-05-03; only the NuGet release remains.

Phase 2 prerequisites (must be settled before the NuGet release ships):
- *Reconcile version numbers across artifacts* — needs a deeper conversation before a decision; **currently deferred**, do not implement until discussed. The PS module is now at `4.6.1`; the NuGet is at `2.6.0`. See [`futures/releases.md`](futures/releases.md#reconcile-version-numbers-across-artifacts).
- ✅ *Investigate flakiness from leftover Visio processes* — done in Phase 1 (orphan-leak fix); resolution detail in [`COMPLETED.md`](COMPLETED.md#investigate-flakiness-from-leftover-visio-processes).

## Phase 3 — Modernization *(in progress)*

Phase 3 items completed (so far):
- ✅ *Migrate from `packages.config` to `PackageReference`* — all 11 csprojs converted; Central Package Management; dev-pack install requirement gone via `Microsoft.NETFramework.ReferenceAssemblies` packages. Detail in [`COMPLETED.md`](COMPLETED.md#migrate-from-packagesconfig-to-packagereference).
- ✅ *Modernize SDK-style csproj* — all 11 csprojs converted to SDK-style; net -1,322 lines across the three sub-passes (libraries, tests, exes); MSB3270 mismatch + filename-casing fix + 7-year-old dead code surfaced and removed as side benefits. Detail in [`COMPLETED.md`](COMPLETED.md#modernize-sdk-style-csproj).
- ✅ *Test-discovery linter* (`MSTest.Analyzers` + MSTEST0030 enforcement) and *per-project test READMEs / `docs/TESTING.md`* — closed most of the *General cleanup of the test projects* entry from the Tests section below; only the *Coverage gaps* angle remains. Detail in [`COMPLETED.md`](COMPLETED.md#test-discovery-linter-msttestanalyzers--mstest0030-enforcement).

Phase 3 items still pending:
- *Move development to Visual Studio 2026* — gated on the TFM bump. See [`futures/build-and-code.md`](futures/build-and-code.md#move-development-to-visual-studio-2026).
- *Consolidate target frameworks* — step 2 (4.5 → 4.7.2). **Deferred until 2026-10-13** when Windows 10 LTSB 2016 leaves Extended Support. See [`futures/build-and-code.md`](futures/build-and-code.md#consolidate-target-frameworks).
- *Consider migrating off Visio 2010 PIA* — see [`futures/build-and-code.md`](futures/build-and-code.md#consider-migrating-off-visio-2010-pia).
- *Decide whether to move to .NET 6/8 (out of .NET Framework)* — see [`futures/build-and-code.md`](futures/build-and-code.md#decide-whether-to-move-to-net-68-out-of-net-framework).
- *Automate releases via GitHub CI — NuGet + PowerShell Gallery* — partially in progress (binary-only GitHub Releases done; PSGallery + nuget.org publish steps pending). See [`futures/releases.md`](futures/releases.md#automate-releases-via-github-ci-in-progress).
- *Decide where docs live long-term* — see [`futures/docs.md`](futures/docs.md#decide-where-docs-live-long-term).
