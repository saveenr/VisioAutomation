
# The Developer command group is removed

**Status:** Accepted (2026-11-20)

## Context

`client.Developer` (`VisioScripting.Commands.DeveloperCommands`) was a facade group that let the library document itself. It drew Visio diagrams of the library's own namespaces and classes (`DrawNamespaces`, `DrawNamespacesAndClasses`), drew documents of the scripting API and of the Visio interop enums (`DrawScriptingDocumentation`, `DrawInteropEnumDocumentation`), and exposed the enum data behind them (`GetInteropEnums`, `GetInteropEnum`, `GetEnum`, `GetTypes`). It was 11 public methods, about 500 lines, backed by `InteropHelper` and the public `EnumType` and `EnumValue` models.

Who used it:

- **No cmdlet.** Nothing in `VisioPowerShell` calls it, and `Visio.psd1` exports no matching cmdlet.
- **Four samples** in `VSamples/Samples/Developer/` and **three tests** in `VTest.Scripting`.
- **No known outside callers.** Four of its methods were already listed as dead (no callers) in [#183](https://github.com/saveenr/VisioAutomation/issues/183).
- `DrawScriptingDocumentation` rendered an empty document (see [`form-page-model-internal.md`](form-page-model-internal.md)).

## Decision

Remove it: `DeveloperCommands`, the `Client.Developer` property, `InteropHelper`, `VisioScripting.Models.EnumType` and `EnumValue` (public, and used nowhere else), the five `VSamples` developer samples and their registrations, and the three developer tests. `Client` now has 24 command groups.

Left in place: `ReflectionHelper` (still used by `Command.cs`) and the form page model under `VisioScripting/Helpers/Forms/`. Its two `client.Developer` callers are gone, so the only remaining consumer is the `MonitorResolutions` sample. What to do with it is a separate decision.

The group was not marked `[Obsolete]` for a release first. Like the earlier internalizations (see [`visioscripting-public-api.md`](visioscripting-public-api.md)), no outside callers are known, and a deprecation window for something nobody is known to use only delays the cleanup.

## Why

- **The job it did is better done elsewhere now.** Diagrams and documents of a code base's structure are something an AI coding agent or a reflection tool can produce directly from the source, on demand, in whatever shape is wanted. A built-in generator that only draws the library's own API is maintenance cost with no end user.
- **It shrinks the public contract.** `Client` and every `*Commands` class are public-stable, so each method is a commitment. Removing a whole group removes 11 methods and two model types from that commitment.
- **It was unreachable from PowerShell,** the library's main consumer.

## Consequences

- **Breaking for anyone who called `client.Developer.*`, or used `EnumType` or `EnumValue`.** The NuGet changelog records it under `[Unreleased]` / Removed. Version number and release notes need the same decision as the other internalizations (see `CLAUDE.md`).
- **No change to the PowerShell module.** No cmdlet is affected, so the module changelog needs no entry.
- **The GitBook has a page for this group** (one page per `*Commands` class under `visio-scripting/`). It needs removing, and the overview that lists the groups needs its count and list updated, before the release that ships this. Those pages live in a separate repository and are not changed here.
- **Part of [#183](https://github.com/saveenr/VisioAutomation/issues/183) is now done.** Its `DeveloperCommands` entries no longer apply.
- **The Visio interop enum documents are gone.** If one is wanted again, generate it from the interop assembly outside the library.

## Cross-references

- [#183](https://github.com/saveenr/VisioAutomation/issues/183) (dead-method removal), [#184](https://github.com/saveenr/VisioAutomation/issues/184) (dead-method stance), [#220](https://github.com/saveenr/VisioAutomation/issues/220) (the form page model).
- [`visioscripting-public-api.md`](visioscripting-public-api.md): the `Client` contract this changes.
- [`form-page-model-internal.md`](form-page-model-internal.md): partly superseded by this record.

## Reconsider when

- Users ask for built-in diagrams or documents of the API or the Visio interop types, or an outside caller of `client.Developer` turns up.
