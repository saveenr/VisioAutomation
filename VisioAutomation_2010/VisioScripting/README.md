# VisioScripting

A scripting-friendly facade over `VisioAutomation` (core) and `VisioAutomation.Models`. Organizes operations into ~25 verb-noun **command groups** hung off a single [`Client`](Client.cs).

Depends on `VisioAutomation` and `VisioAutomation.Models`. Used by `VisioPowerShell`. For the layered architecture see [`docs/ARCHITECTURE.md`](../../docs/ARCHITECTURE.md).

Built into the [`VisioAutomation2010`](https://www.nuget.org/packages/VisioAutomation2010/) NuGet package alongside `VisioAutomation` and `VisioAutomation.Models`. Release notes: [`NuGet/CHANGELOG.md`](../../NuGet/CHANGELOG.md).

## Typical usage

```csharp
var client = new VisioScripting.Client(visioApp);
client.Document.NewDocument();
client.Draw.DrawRectangle(TargetPage.Auto, new Rectangle(0, 0, 4, 2));
client.Text.SetText(TargetShapes.Auto, "Hello");
```

## Key types

- [`Client`](Client.cs) — entry point, constructed with an `IVisio.Application`. Exposes command groups as properties (`Application`, `Document`, `Page`, `Selection`, `Draw`, `Text`, `Arrange`, `Connection`, `ShapeSheet`, `Layer`, `Grouping`, `Master`, `CustomProperty`, `Hyperlink`, `Control`, …).
- [`ClientContext`](ClientContext.cs) — abstract output sink (`WriteDebug` / `WriteUser` / `WriteError` / `WriteVerbose` / `WriteWarning`). Subclass to redirect logging.
- [`DefaultClientContext`](DefaultClientContext.cs) — concrete `ClientContext` that writes to `Console`.
- [`CommandTarget`](CommandTarget.cs) + [`CommandTargetFlags`](CommandTargetFlags.cs) — preconditions wrapper. A command declares it needs (e.g.) an active page; `CommandTarget` validates and resolves that state up front.
- `Target*` family — deferred-resolution wrappers ([`TargetDocument`](TargetDocument.cs), [`TargetPage`](TargetPage.cs), [`TargetShapes`](TargetShapes.cs), [`TargetSelection`](TargetSelection.cs), [`TargetWindow`](TargetWindow.cs), `TargetPages`, `TargetDocuments`, `TargetObject`, `TargetObjects`). `TargetPage.Auto` means *"use the active page when this command runs"* — keeps callers from having to fetch and pass COM objects explicitly.

## Folder layout

- `Commands/` — verb-noun command groups, one file per group. `ApplicationCommands`, `DocumentCommands`, `PageCommands`, `DrawCommands`, `SelectionCommands`, `TextCommands`, `ShapeSheetCommands`, `ArrangeCommands`, `ConnectionCommands`, `ConnectionPointCommands`, `ContainerrCommands`, `ControlCommands`, `CustomPropertyCommands`, `DeveloperCommands`, `ExportCommands`, `GroupingCommands`, `HyperlinkCommands`, `LayerCommands`, `LockCommands`, `MasterCommands`, `ModelCommands`, `OutputCommands`, `UndoCommands`, `UserDefinedCellCommands`, `ViewCommands`. Plus `Command.cs`/`CommandParameter.cs`/`CommandSet.cs` (base types).
- `Models/` — small enums and value types used as command parameters and return values (`AlignmentHorizontal`, `AlignmentVertical`, `Axis`, `ConnectionPointType`, `PageDimensions`, `PageOrientation`, `SelectionOperation`, `ShapeDimensions`, …).
- `Helpers/` — internal helpers shared across command groups (`ArrangeHelper`, `InteropHelper`, `ReflectionHelper`, `SelectionHelper`, `TextHelper`, `WildcardHelper`).
- `Loaders/` — turn-key loaders for higher-level documents (`DirectedGraphDocumentLoader`, `OrgChartDocumentLoader`).
- `Extensions/` — extension methods used internally (`XmlLinqExtensions`).

## See also

- [`docs/ARCHITECTURE.md`](../../docs/ARCHITECTURE.md) — solution-wide architecture and dependencies
- [`docs/GLOSSARY.md`](../../docs/GLOSSARY.md) — Visio + codebase terminology
- [`docs/BUILDING.md`](../../docs/BUILDING.md) — how to build, test, install
- [`NuGet/CHANGELOG.md`](../../NuGet/CHANGELOG.md) — release notes for the bundled NuGet package
