
# The Box layout is removed

**Status:** Accepted (2026-11-20; see [#218](https://github.com/saveenr/VisioAutomation/issues/218) for the discussion)

## Context

`VisioAutomation.Models.Layouts.Box` (`BoxLayout`, `Container`, `Box`, `Node`, `Direction`, `AlignmentHorizontal`, `AlignmentVertical`; about 380 lines) packed nested rectangles: a `Container` had a direction, padding, child spacing and a minimum size, and its children were `Box` leaves or further containers, each with an alignment inside its parent. `BoxLayout.PerformLayout()` computed the rectangles. It had no Visio dependency and drew nothing.

It was public API in the `VisioAutomation2010` NuGet package, but it was not reachable from anywhere:

- **No entry point.** No `client.Model` method, no cmdlet and no DOM renderer drew it. A caller had to walk the computed rectangles and draw them by hand.
- **No production caller** in any project of this repository.
- **Used only by samples and tests:** three `VSamples` samples (`BoxLayoutSimple`, `BoxLayoutTwoLevelGrouping`, `CompareFonts`, which drew through a helper that called `DrawRectangle` for each rectangle), the four `BoxLayoutTests`, and `DomTextTests`, which only borrowed `AlignmentHorizontal.Left`, `Center` and `Right` as the numbers 0, 1 and 2 for Visio paragraph alignment.
- **Thin evidence of users.** A placement bug in its `RightToLeft` direction sat unnoticed for years and was fixed in 3.1.0. A GitHub code search for the namespace finds only this repository and its GitBook docs. That search cannot see NuGet consumers or private code.

The .NET GitBook had already moved it out of "Layout models" into a "Box geometry" page that said it only computes rectangles and that nothing else uses it. [#218](https://github.com/saveenr/VisioAutomation/issues/218) listed the options: keep it as a documented helper, give it a renderer, deprecate and remove it, or move it to another namespace.

## Decision

Remove it. The seven types, the three samples and their sample helper (`Util.cs`), the sample registrations, and `BoxLayoutTests` are deleted. `DomTextTests` now uses named constants (0, 1 and 2) for Visio's paragraph alignment values instead of the Box enum.

It was not marked `[Obsolete]` for a release first. That matches the form page model and the Developer group ([`form-page-model-internal.md`](form-page-model-internal.md), [`developer-commands-removed.md`](developer-commands-removed.md)): no outside callers are known, and a deprecation window for a type nobody is known to use only delays the cleanup.

## Why

- **It was public surface with no way to use it.** Under the public-API approach, each public type is a commitment. This one could not produce a drawing without hand-written code, and nothing in the library needed it.
- **It was easy to mistake for a layout model,** and `Container` in it is easy to confuse with the Container layout ([#221](https://github.com/saveenr/VisioAutomation/issues/221)).
- **Keeping it would mean building a renderer for a feature with no known demand.** If one turns up, the packing idea is small enough to bring back.

## Consequences

- **Breaking for anyone who used the Box types directly.** The NuGet changelog records it under `[Unreleased]` / Removed. Version number and release notes need the same decision as the other removals under `[Unreleased]` (see `CLAUDE.md`).
- **The GitBook needs a change before the release that ships this.** Delete `models/box-geometry.md` and its `SUMMARY.md` entry, the Box rows on `models/introduction.md`, and the Box mention in the Layout models overview. Those pages live in a separate repository.
- **Three samples are gone.** `CompareFonts`, a font comparison, was a Box demo and went with it. If it is wanted, it can be rewritten against the DOM or the Grid layout.
- **`DomTextTests` no longer depends on a layout enum.**

## Cross-references

- [#218](https://github.com/saveenr/VisioAutomation/issues/218) (this decision), [#221](https://github.com/saveenr/VisioAutomation/issues/221) (the Container layout, still open).
- [`form-page-model-internal.md`](form-page-model-internal.md), [`developer-commands-removed.md`](developer-commands-removed.md): the earlier removals of public types with no real caller.

## Reconsider when

- A user asks for nested rectangle packing, or there is evidence of outside callers of the Box types.
- The Container layout decision ([#221](https://github.com/saveenr/VisioAutomation/issues/221)) wants a packing engine of this kind.
