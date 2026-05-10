# Testing

Design and conventions of the test suite. For *running* tests (commands, prerequisites, IDE flow) see [BUILDING.md](BUILDING.md). For per-project breakdowns see each project's own README.

## The four test projects

All under `VisioAutomation_2010/`:

| Project | Tests | Library under test | README |
|---|---:|---|---|
| `VTest` | 94 | `VisioAutomation` (core) | [VTest/README.md](../VisioAutomation_2010/VTest/README.md) |
| `VTest.Models` | 45 | `VisioAutomation.Models` (DOM, layouts) | [VTest.Models/README.md](../VisioAutomation_2010/VTest.Models/README.md) |
| `VTest.Scripting` | 34 | `VisioScripting` (high-level facade) | [VTest.Scripting/README.md](../VisioAutomation_2010/VTest.Scripting/README.md) |
| `VTest.PowerShell` | 4 | `VisioPowerShell` (cmdlets) | [VTest.PowerShell/README.md](../VisioAutomation_2010/VTest.PowerShell/README.md) |

177 tests total. Counts as of 2026-05-04.

## Framework: MSTest 4.x

Tests use **MSTest 4.x** (`MSTest.TestAdapter` + `MSTest.TestFramework` 4.2.2). The MSTest 4.x upgrade landed in Phase 1 of the 2026 refresh; the suite was previously on the MSTest beta. See [futures/tests.md](futures/tests.md#evaluate-modern-testing-stack-options) → *Evaluate modern testing-stack options* for the survey of alternatives (Verify, Shouldly, FsCheck, xUnit) considered and what we picked / deferred / rejected.

`MSTest.Analyzers` 4.2.2 is wired into all four test projects. Severity defaults are kept for most rules; **MSTEST0030 is promoted to `error`** in [`VisioAutomation_2010/.editorconfig`](../VisioAutomation_2010/.editorconfig). See *Quality gates* below.

## Design constraints

These three are load-bearing — every other choice in the test infrastructure follows from them.

### 1. Tests run against a real Visio install

There is no mock or fake Visio. Tests instantiate `Microsoft.Office.Interop.Visio.Application`, manipulate real shapes, and verify behavior by reading back actual ShapeSheet values.

**Full decision record:** [`decisions/tests-need-visio.md`](decisions/tests-need-visio.md) covers the *why* (mocks would have to reproduce Visio's undocumented quirks), the consequences (no `dotnet test` that runs anywhere; CI today is build-only; self-hosted Windows runner gated for any future test-CI per [`futures/build-and-code.md`](futures/build-and-code.md#run-tests-in-ci)), and the de-facto no-Visio test bucket that's emerging (`ManifestTests`, `XmlErrorLogTests`).

### 2. Tests run sequentially, not in parallel

MSTest's parallel execution is not enabled. Visio doesn't tolerate concurrent COM clients well, and the bottleneck for this suite is Visio's cold-start time, not test-runner overhead.

### 3. One Visio process per test host (singleton)

Each test project shares a single `Visio.Application` instance across all its tests via `VTest.Framework.VTestAppRef`. Tests obtain it through `this.GetVisioApplication()` (inherited from `Framework.VTest`) or via `GetScriptingClient()`. Reusing the instance avoids paying Visio's cold-start cost per test.

## Shared infrastructure

Three of the four test projects (`VTest`, `VTest.Models`, `VTest.Scripting`) inherit from a common base class and share lifecycle infrastructure. `VTest.PowerShell` is the exception — it tests through a PowerShell runspace and has its own pattern.

### `Framework.VTest` base class

Lives in `VTest/Framework/VTest.cs`. Marked `[TestClass]` itself but contains no test methods. Provides:

- `GetVisioApplication()` — returns the per-testhost singleton.
- `GetNewPage()` / `GetNewPage(string suffix)` / `GetNewPage(Size)` — creates a fresh page; uses the calling method name as the page name (via `StackFrame`).
- `GetNewDoc()` — creates a new document and tags it with the calling method's name.
- `GetScriptingClient()` — returns a `VisioScripting.Client` wrapping the singleton.
- `GetSize(IVisio.Shape)` / `SetPageSize(...)` / `GetPageSize(...)` — geometry helpers.
- `get_datafile_content(name)` — loads test data files from the build output directory.

### `Framework.VTestAppRef` singleton

`VTest/Framework/VTestAppRef.cs`. Per-testhost field of type `IVisio.Application`. `GetVisioApplication()` lazily creates the instance, recreates it on `COMException` (i.e., if Visio was closed externally between tests). `QuitVisioApplication()` closes all open documents and quits Visio — called from each project's `[AssemblyCleanup]`.

### `[AssemblyCleanup]` orphan-prevention

Each test project carries its own `AssemblyHooks.cs` with an `[AssemblyCleanup]` that calls `VTestAppRef.QuitVisioApplication()`. **Don't try to share this via the base class** — `[AssemblyCleanup]` is per-assembly and not inherited. Phase 1 commit `9a592a9d` added these hooks after discovering each testhost was leaking its singleton on exit (4 orphans per clean run, ~945 MB).

### Data files

Test fixtures (`VTest/datafiles/*`) are tagged `<Content Include="..." CopyToOutputDirectory="Always" />` in the csproj. They get copied alongside the test DLL automatically.

**Don't add `[DeploymentItem]` attributes.** Phase 1 commit `5cbf11cd` removed them — they were redundant given the `CopyToOutputDirectory` flag, and their only effect was triggering VS Test Explorer's deployment-mode behavior, which in turn dropped runtime dependencies on the floor.

### `VTest.PowerShell`: cmdlet-binding tests via `InvokeScript` / `InvokeScriptStrict`

`VTest.PowerShell` doesn't share the `Framework.VTest` base class; it tests cmdlets via a real PowerShell runspace hosted by [`VisioPSSession`](../VisioAutomation_2010/VTest.PowerShell/VisioPSSession.cs). Two paths are available:

- **`Cmd_*` helpers** (e.g. `Cmd_New_VisioDocument`) — instantiate a cmdlet object in C# and call `cmd.Invoke()` directly. Bypasses PowerShell's parameter binder. Convenient for setup, but **wrong for tests of binding behavior**.
- **`InvokeScript<T>` / `InvokeScriptStrict<T>`** — execute a PowerShell script through the runspace, exercising the real binder. Required for any test of positional binding, switch parameters, parameter sets, or pipeline binding.

#### When to use which `InvokeScript` variant

`InvokeScriptStrict<T>` is `InvokeScript<T>` with `$ErrorActionPreference = 'Stop'` prepended. PowerShell catches cmdlet-thrown exceptions and writes them to the error stream by default; without `'Stop'`, a thrown exception **does not propagate** to the caller, so a `try`/`catch` around `InvokeScript<T>` never sees it.

- Use **`InvokeScriptStrict<T>`** for tests that expect the cmdlet to throw (so the test can catch the propagated exception), or that want any unexpected error from the cmdlet to surface as a test failure rather than be silently dropped on the error stream. **This is the right default for cmdlet-binding tests** ([`CmdletBindingTests.cs`](../VisioAutomation_2010/VTest.PowerShell/CmdletBindingTests.cs) is the canonical example).
- Use plain **`InvokeScript<T>`** when the cmdlet is allowed to write non-fatal records to the error stream and the test only cares about the success-path return value.

The longer-term cleanup (migrating cmdlets from raw `throw` to `ThrowTerminatingError(ErrorRecord)`, which always propagates regardless of `$ErrorActionPreference`) is tracked in [#191](https://github.com/saveenr/VisioAutomation/issues/191). Until that lands, `InvokeScriptStrict<T>` is the correct workaround for binding-test exception assertions.

## Quality gates

### MSTEST0030 enforced as error

The analyzer rule **MSTEST0030** ("Type containing `[TestMethod]` should be marked with `[TestClass]`") is promoted from the MSTest.Analyzers default warning to `error` in [`VisioAutomation_2010/.editorconfig`](../VisioAutomation_2010/.editorconfig).

**Why.** Phase 1 commit `b77a99f0` discovered that 14 test methods (~8% of the suite) had been silently skipped for years because seven test classes deriving from `Framework.VTest` lacked the `[MUT.TestClass]` attribute on the class declaration. MSTest 4.x doesn't inherit `[TestClass]` from a base class, and the build emitted no warning. Promoting MSTEST0030 to error means the regression cannot recur silently — the build fails immediately if any test class with `[TestMethod]` members lacks `[TestClass]`. Especially important once release CI lands; a release pipeline that gates on a test suite which silently shrinks is worse than no pipeline.

The rule excludes `abstract` classes by design, so a `[TestClass]`-less abstract base would not trip it (we don't use abstract base classes in this suite, but it's the documented escape hatch).

### Other MSTest analyzer rules

The other MSTEST00xx rules ship at default severity (mostly warnings). They've found one real issue so far (an `Assert.AreEqual` argument-order swap, fixed alongside the analyzer wiring). If you bump a rule's severity, document the reason in `.editorconfig` so the choice is explainable later.

## Naming conventions

### Test methods

New tests should follow `Subject_Scenario_ExpectedOutcome`:

- `Loader_ConnectorType_DefaultsToCurvedWhenAttributeMissing`
- `DirectedGraph_NodeSizeIsHonored_WhenCellsAlsoSet`
- `Application_UndoScope_NestedInner`
- `GetVisioShape_NoArgs_ReturnsAllShapesOnPage`
- `SetVisioUserDefinedCell_EncodesValueAndPrompt`

Each segment carries information: `Subject` says *what* is under test, `Scenario` says *under what conditions*, `ExpectedOutcome` says *what should happen*. A failure message that includes only the test name should be enough to know roughly what broke.

### Anti-patterns to avoid

- **Numbered suffixes without a scenario hint.** `Path_TestTransitiveClosure0/1/2/3/4`, `Container_Diagram1/2`, `Dom_ConnectShapes2`, solitary `_1` of nothing (`XmlErrorLog_Load_Visio2010_1`). Numbers are fine *after* a descriptive scenario (`OrgChart_FiveNodes` is OK), but `_2` alone is uninformative.
- **"Scenarios" / "Scenario" kitchen-sinks.** `Scripting_Hyperlinks_Scenarios` admits the test does many things in one method — when it fails, the failure tells you nothing about which scenario broke. Split into separately-named tests; one assertion-focused method per scenario.
- **Redundant `Test_` infix.** `Scripting_Test_Resize_Application_Window1` is two redundant words: `Scripting_` is the prefix, every method is a test. Drop both — `Application_ResizeWindow_TaskbarHidden` says more in less.
- **Vague single-word names.** `Basics`, `QueryPage`, `Connect1` — open the body to know what's being verified.
- **Inconsistent prefix within a single file.** Pick one prefix per file (the area or class under test) and stick to it. Mixing `VSD_Load_*` with `XmlErrorLog_Load_*` in the same class is purely historical drift.

### Test files

`<Subject>Tests.cs`. PascalCase, plural `Tests`, no underscores between subject and `Tests`:

- ✅ `BoundingBoxHelperTests.cs`, `ApplicationHelperTests.cs`, `DOM_Tests.cs` (the `DOM_` is the subject; the tail is bare `Tests`).
- ❌ `Path_Test.cs` (singular), `ConnectionPoint_Tests.cs` (extra underscore), `Dom_Text_Tests.cs` (extra underscore).

### Test classes

Match the file name. Marked `[TestClass]` directly on the concrete class — MSTest 4.x doesn't inherit it from `Framework.VTest`, and MSTEST0030 (promoted to error) catches violations at build time. See *Quality gates* above.

## Running

See [BUILDING.md](BUILDING.md) for IDE flow, `vstest.console.exe` invocation, and the dev-pack install requirement. Quick reminders:

- Visio must be installed locally.
- Tests are sequential — don't expect parallel speedup.
- A clean run of all 177 tests should leave **zero** Visio orphan processes. If you see Visio in Task Manager after a green run, that's a regression in the assembly-cleanup wiring (commit `9a592a9d` is the canonical fix to look at).
- Interrupted runs (Ctrl-C, debugger detach) can leave orphans behind; close them in Task Manager before re-running, or successive runs may hit file locks on stencils / templates.

## Known gotchas

- **`[TestClass]` doesn't inherit.** Covered above; MSTEST0030 catches this now.
- **`[AssemblyCleanup]` doesn't inherit either.** Each test project needs its own `AssemblyHooks.cs`. (Same MSTest-attribute-inheritance gotcha as above.)
- **Orgchart template filename changed in Visio 2013** (`.vst` → `.vstx`). Tests opening the orgchart stencil version-guard on `app.Version` — see commit `da9bba0a` and `OrgChartStyling.cs`. New tests touching that stencil should follow the same pattern.
