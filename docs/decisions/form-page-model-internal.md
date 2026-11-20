
# The form page model is internal to VisioScripting

**Status:** Accepted (2026-11-20; see [#220](https://github.com/saveenr/VisioAutomation/issues/220) for the discussion). Superseded the same day by [`developer-commands-removed.md`](developer-commands-removed.md): the two `client.Developer` callers described below were removed, and the form types were then deleted along with their last consumer, the `MonitorResolutions` sample. The reasoning below is kept as history.

## Context

`FormDocument`, `FormPage`, `TextBlock`, `InteractiveRenderer` and `PageMargin` generated printable, document-style Visio pages (a title and a body per page, in a new document). They lived in `VisioAutomation.Models.Documents.Forms` and were public API in the `VisioAutomation2010` NuGet package. The .NET GitBook presented them as one of two "Document models", next to the org chart model.

In practice they were a small generator used by the library's own developer tooling:

- One working internal caller: `client.Developer.DrawInteropEnumDocumentation()`, behind the `DocumentVisioInterop` sample.
- One internal caller that renders an empty document: `client.Developer.DrawScriptingDocumentation()`.
- One sample, `VSamples/Samples/Misc/MonitorResolutions.cs`, which uses the renderer directly.
- No end-user entry point: no `client.Model` method, no cmdlet, no XML loader. Nothing in `VisioPowerShell` referenced it.
- No tests of its own.

[#220](https://github.com/saveenr/VisioAutomation/issues/220) listed four options: keep it public as is, keep it public but stop presenting it as a peer model, make it internal to VisioScripting's developer tooling, or give it a real entry point so it earns its place.

## Decision

Make it internal (the third option). The six files moved from `VisioAutomation.Models/Documents/Forms/` to `VisioScripting/Helpers/Forms/`, namespace `VisioScripting.Helpers.Forms`, and every type is `internal`. `VisioScripting` grants `InternalsVisibleTo("VSamples")` so the one sample keeps working.

The types were **not** marked `[Obsolete]` for a release first. That matches the earlier internalization of the `Helpers`, `Loaders` and `CommandTarget` plumbing ([`visioscripting-public-api.md`](visioscripting-public-api.md)), which also went straight to `internal` for types with no known outside callers. Here the only evidence of that is the code search described below, so a deprecation window would have been the more cautious path; it was skipped because a window for a type nobody is known to use only delays the cleanup.

## Why

- **It fits how it is used.** Its only real consumer is VisioScripting's developer tooling, so it now lives there and `VisioAutomation.Models` no longer carries code that no model user would reach for.
- **It stops overstating the model.** A public "Document model" with no entry point and no tests suggested a supported feature that did not exist.
- **The evidence of outside use is thin.** A GitHub code search for the namespace finds only this repository and its GitBook docs. That search cannot see NuGet consumers or private code, so the risk is not zero, and it is accepted.

## Consequences

- **Breaking for anyone who used the types directly.** The NuGet changelog records it under `[Unreleased]` / Removed. Per [`CLAUDE.md`](../../CLAUDE.md), resolve the compatibility implications of everything internalized under `[Unreleased]` before publishing; this one needs the same call (version number and release notes) as the other internalizations.
- **The GitBook needs a change before the release that ships this.** `models/forms.md` describes types that are no longer public, and `models/documents.md` lists them. Those pages live in a separate repository, so the change is made there. How to stage docs for an unreleased change is still open ([#238](https://github.com/saveenr/VisioAutomation/issues/238)).
- **`DrawScriptingDocumentation` still renders an empty document.** This decision does not change that. Fixing or removing it is a separate item.
- **The form types are still untested directly.** One smoke test now covers the working caller, `DevTests.DrawInteropEnumDocumentation_NewDocument_HasPagesWithShapes`.
- **The sample reaches into internals.** `MonitorResolutions` uses the renderer through `InternalsVisibleTo`. If more samples need it, or it starts to look like a public example, move the sample onto a public API instead.

## Cross-references

- [#220](https://github.com/saveenr/VisioAutomation/issues/220) (this decision), [#218](https://github.com/saveenr/VisioAutomation/issues/218) (the Box layout raises the same question), [#182](https://github.com/saveenr/VisioAutomation/issues/182) (the earlier internalizations).
- [`visioscripting-public-api.md`](visioscripting-public-api.md): what is public-stable in VisioScripting, and why plumbing is internal.
- Code: `VisioAutomation_2010/VisioScripting/Helpers/Forms/`.

## Reconsider when

- A user asks to generate document-style Visio pages from their own code, or there is evidence of external callers of the old public types.
- The Box layout decision ([#218](https://github.com/saveenr/VisioAutomation/issues/218)) goes the other way and the project wants a family of small public utility models.
