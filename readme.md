# VisioAutomation

[![Build](https://github.com/saveenr/VisioAutomation/actions/workflows/build.yml/badge.svg)](https://github.com/saveenr/VisioAutomation/actions/workflows/build.yml)
[![NuGet](https://img.shields.io/nuget/v/VisioAutomation2010.svg)](https://www.nuget.org/packages/VisioAutomation2010/)
[![PowerShell Gallery](https://img.shields.io/powershellgallery/v/Visio.svg)](https://www.powershellgallery.com/packages/Visio)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE.txt)

A set of .NET libraries that make it easier to control [Microsoft Visio](https://www.microsoft.com/microsoft-365/visio/flowchart-software) from .NET languages. In addition to simplifying common tasks, they make it easier to build your own Visio add-ins, automation tools, and scripts.

The project ships two artifacts:

| Artifact | What it is | Install |
|---|---|---|
| [`VisioAutomation2010`](https://www.nuget.org/packages/VisioAutomation2010/) | The .NET library | `Install-Package VisioAutomation2010` |
| [`Visio`](https://www.powershellgallery.com/packages/Visio) | A PowerShell module that wraps the library | `Install-Module Visio` |

## Quick example

C# (using the high-level scripting facade):

```csharp
var visio  = new Microsoft.Office.Interop.Visio.Application();
var client = new VisioScripting.Client(visio);

client.Document.NewDocument();
client.Draw.DrawRectangle(VisioScripting.TargetPage.Auto,
                          new VisioAutomation.Core.Rectangle(0, 0, 4, 2));
client.Text.SetText(VisioScripting.TargetShapes.Auto, "Hello, Visio!");
```

PowerShell:

```powershell
Import-Module Visio

New-VisioApplication
New-VisioDocument
New-VisioShape -Rectangle -BoundingBox (New-VisioRectangle 0 0 4 2)
```

## Documentation

User guides (recommended first stop):

- [VisioAutomation user guide](https://saveenr.gitbook.io/visioautomation/) (gitbook)
- [Visio PowerShell user guide](https://saveenr.gitbook.io/visiopowershell/) (gitbook)

Developer / architecture docs in this repo:

- [`docs/OVERVIEW.md`](docs/OVERVIEW.md) — index of all developer docs
- [`docs/ARCHITECTURE.md`](docs/ARCHITECTURE.md) — what each project does and how they fit together
- [`docs/BUILDING.md`](docs/BUILDING.md) — build, test, install
- [`docs/GLOSSARY.md`](docs/GLOSSARY.md) — Visio and codebase terminology
- [`docs/FUTURES.md`](docs/FUTURES.md) — staged backlog of cleanup / modernization work

Release notes:

- [NuGet CHANGELOG](NuGet/CHANGELOG.md)
- [PowerShell module CHANGELOG](VisioAutomation_2010/VisioPowerShell/CHANGELOG.md)

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md). Development happens on `master`.

## License

[MIT](LICENSE.txt). Copyright (c) Saveen Reddy.
