# `VisioScripting.Client` is a hybrid public API

**Status:** Accepted (codified 2026-05-09; see [#156](https://github.com/saveenr/VisioAutomation/issues/156) for the decision thread)

## Context

`VisioScripting` is the higher-level facade between the low-level `VisioAutomation` library and the `VisioPowerShell` cmdlets. The entry point is `VisioScripting.Client`, which exposes 25 group properties (`Application`, `Arrange`, `Connection`, `ConnectionPoint`, `Container`, `Control`, `CustomProperty`, `Developer`, `Document`, `Draw`, `Export`, `Grouping`, `Hyperlink`, `Layer`, `Lock`, `Master`, `Model`, `Output`, `Page`, `Selection`, `ShapeSheet`, `Text`, `Undo`, `UserDefinedCell`, `View`). Each property returns a `*Commands` instance whose public methods do the actual work.

The repo's [`readme.md`](../../readme.md) opens its C# quick-start with `new VisioScripting.Client(app)`, so direct .NET consumers are not hypothetical &mdash; but the type was previously undocumented on the gitbook and treated as project-internal in practice. We had to choose:

1. **Fully public.** Document the entire surface (every method on every `*Commands`, plus `Helpers/`, `Loaders/`, `CommandTarget`, etc.) and treat all of it as a public contract under SemVer.
2. **Fully internal.** Mark the namespace as not-for-direct-use, fix the readme quick-start to use `VisioAutomation` directly, close out [#131](https://github.com/saveenr/VisioAutomation/issues/131) won't-fix.
3. **Hybrid.** Public-stable for the `Client` facade itself; implementation detail for the plumbing.

The pre-decision usage audit found that `Client.<Group>.<Method>` accounts for ~76% of method usage across cmdlets and tests. Two cmdlets bypassed the facade (the `Loaders` reach-in and the `*Dimensions.Get_*` reach-in), but both were small and migrating them was straightforward.

## Decision

Hybrid (option 3).

## The contract

### Public-stable (documented; breaking changes treated as breaking)

- `VisioScripting.Client` &mdash; the type, both constructors, and the 25 group properties.
- The public method signatures on each `*Commands` class. Adding methods is non-breaking; renaming or removing methods is breaking.
- `Target*` types at the namespace root (`TargetDocument`, `TargetDocuments`, `TargetPage`, `TargetPages`, `TargetSelection`, `TargetShapes`, `TargetWindow`). Forced &mdash; they appear in `*Commands` method signatures.
- `VisioScripting.Models.*` types referenced from public signatures (enums like `PageOrientation`, `ZoomToObject`, `ShapeSelectionOperation`; data carriers like `PageDimensions`, `ShapeDimensions`, `ShapeSheetReader`, `ShapeSheetWriter`).
- `ClientContext` and `DefaultClientContext`. The `Client(app, ClientContext)` overload exists so an embedding host (today `VisioPowerShell`'s [`VisioPsClientContext`](../../VisioAutomation_2010/VisioPowerShell/VisioPsClientContext.cs)) can plug in its own `Output` plumbing; subclassing has to remain a supported pattern.

### Internal-mutable (not part of contract; free to change)

- The `*Commands` classes as **constructible types**. The classes themselves are necessarily `public` (they appear as the return type of `Client.Page`, `Client.ShapeSheet`, etc.), but their constructors are `internal`. Consumers obtain instances via `Client.<Group>`, never `new`.
- `VisioScripting.Helpers/*` &mdash; `WildcardHelper`, `InteropHelper`, `SelectionHelper`, `ArrangeHelper`, `ReflectionHelper`, `TextHelper`. Pure utilities.
- `VisioScripting.Loaders/*` &mdash; `DirectedGraphDocumentLoader`, `OrgChartDocumentLoader`. Reached only through `Client.Model.LoadDirectedGraphFromXml(...)` / `Client.Model.LoadOrgChartFromXml(...)`.
- `CommandTarget`, `CommandTargetFlags`. Used inside `*Commands` method bodies to validate preconditions; never appear in a public signature.
- `Client.GetCommandTarget(flags)` &mdash; helper for `*Commands` implementations.
- `Models.*` types not appearing in any public signature (today: `DgShapeInfo`, `DgConnectorInfo`, both already `internal`).
- The static factory methods on data-carrier types: `Models.PageDimensions.Get_PageDimensions(...)`, `Models.ShapeDimensions.Get_ShapeDimensions(...)`. The data classes themselves stay `public` (return values); the factories are `internal`.

## Enforcement

Layered, weakest to strongest:

1. **C# `internal` keyword.** The default mechanism. Every type that can be `internal` is `internal`. Same-assembly callers (the `*Commands` classes calling `Helpers/`, `*Dimensions.Get_*`, `CommandTarget`) keep working.
2. **`[InternalsVisibleTo("VTest")]`** on the `VisioScripting` assembly ([`Properties/AssemblyInfo.cs`](../../VisioAutomation_2010/VisioScripting/Properties/AssemblyInfo.cs)). `VTest` exercises `WildcardHelper.GetRegexForWildcardPattern` directly &mdash; the test is testing the helper itself, not its consumers. None of `VisioPowerShell`, `VTest.Models`, `VTest.Scripting`, or `VTest.PowerShell` need internal access today; they consume `VisioScripting` exclusively through public surface. If a future test or cmdlet needs to reach in, add an `InternalsVisibleTo` line for the specific project rather than blanket-granting access.
3. **`[EditorBrowsable(EditorBrowsableState.Never)]` + XML doc comments.** Held in reserve for any type that has to stay `public` for type-system reasons but isn't part of the contract. None today after the Phase B + C work, but the mechanism is documented here so future code knows to use it.

The `Microsoft.CodeAnalysis.PublicApiAnalyzers` package (`PublicAPI.Shipped.txt` / `PublicAPI.Unshipped.txt`) was deferred. It's the strongest enforcement (a build error when the shipped surface drifts unintentionally), but it's another moving part. Revisit if drift becomes a real problem.

## Consequences

### Documentation

- The doc-write under [#131](https://github.com/saveenr/VisioAutomation/issues/131) covers the public surface listed above and only that surface. The pre-decision estimate was ~152 method signatures across the 25 `*Commands` classes; after the Phase B facade additions in [#182](https://github.com/saveenr/VisioAutomation/issues/182) it's ~156. Plus `Client` itself, the 7 `Target*` types, `ClientContext`/`DefaultClientContext`, and the `Models.*` types referenced from public signatures.
- The ~36 dead methods identified in the audit (zero external callers as of 2026-05-09) are part of the locked surface; their removal is deferred to CY27 in [#183](https://github.com/saveenr/VisioAutomation/issues/183) per the project's "improve before audience-reducing changes" guiding principle. The open sub-question of whether to flag them as removal candidates in their gitbook pages is tracked in [#184](https://github.com/saveenr/VisioAutomation/issues/184).

### Code review

- Removing or renaming a public method on a `*Commands` class is a breaking change. Reviewers and PR authors need to recognize this.
- Adding methods is not breaking. Same for adding properties to `Client`.
- Touching anything in `Helpers/`, `Loaders/`, `CommandTarget`, `CommandTargetFlags`, or the `Get_*Dimensions` static methods is not a public-API change.

### Embedding hosts

- A future host besides `VisioPowerShell` (a VS extension, a different shell, an in-process consumer) can subclass `ClientContext` to plug in its own `Output` routing. The subclass-friendly contract on `ClientContext` is part of the public commitment.

## Cross-references

- [#156](https://github.com/saveenr/VisioAutomation/issues/156) &mdash; the decision thread, including the four-question Q1&ndash;Q4 walkthrough that produced this contract.
- [#182](https://github.com/saveenr/VisioAutomation/issues/182) &mdash; pre-lock cleanup (Phases B + C). Phase B added the four facade methods; Phase C applied the enforcement layer that this ADR records.
- [#183](https://github.com/saveenr/VisioAutomation/issues/183) &mdash; CY27 dead-method removal (Phase A). Cites this ADR.
- [#184](https://github.com/saveenr/VisioAutomation/issues/184) &mdash; the open dead-method-stance sub-question.
- [#131](https://github.com/saveenr/VisioAutomation/issues/131) &mdash; the doc-write for the public surface.
- [`CLAUDE.md`](../../CLAUDE.md) &mdash; per-commit conventions section carries the reviewer-facing summary of "what's a breaking change in `VisioScripting`."
- [`VisioAutomation_2010/VisioScripting/README.md`](../../VisioAutomation_2010/VisioScripting/README.md) &mdash; in-source overview of the facade structure.

## Reconsider when

- **A second embedding host appears that needs a different boundary.** Today's only host is `VisioPowerShell`, which lives in this repo. A VS extension or out-of-repo consumer could surface needs that the current line doesn't serve.
- **The dead-surface backlog ([#183](https://github.com/saveenr/VisioAutomation/issues/183)) ships and the line shifts.** Phase A in CY27 will likely shrink the public surface enough to revisit whether the hybrid mechanism is still the right one, or whether a tighter "fully public for what's left" stance becomes affordable.
- **Drift becomes a real problem.** If the contract erodes via inadvertent breaking changes that slip past code review, escalate to the `PublicAPI.Shipped.txt` analyzer (deferred above) or move the off-contract types to a `VisioScripting.Internal.*` sub-namespace as a more visible signal.
