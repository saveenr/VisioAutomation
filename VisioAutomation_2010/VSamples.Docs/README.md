# VSamples.Docs

The C# code blocks from the .NET documentation (the `VisioAutomation_GitBook_Docs` repository), kept here so they are compiled on every build. If a documented call is renamed, removed or never existed, this project stops building.

## Layout

- `Samples\Docs\` has one file per docs page, mirroring the docs path. `models/directed-graph.md` is `Samples\Docs\models\DirectedGraph.cs`, class `ModelsDirectedGraphSamples`.
- Each C# block on a page is one method, `Sample_01`, `Sample_02` and so on, in page order.
- The block's text sits between two markers. The `id` links it to the docs: the page path without `.md`, then `#`, then the block's position among the page's C# blocks.

```csharp
public static void Sample_03(IVisio.Page page)
{
    // <sample id="pages/format#03" doc="pages/format.md" line="41">
    ...the block, exactly as the docs show it...
    // </sample>
}
```

- Blocks are fragments. They use names the page sets up in prose (`page`, `shape`, `client`, a cell record from an earlier block), so those names are method parameters. A block that continues an earlier one on the same page may repeat that block's code above the opening marker, under a "setup" comment. Only the text between the markers is compared with the docs.
- `using` lines from a block are collected at the top of the file. Aliases such as `VA` that a page uses without declaring are added there too.
- Nothing runs these methods. They compile; they are not called. `Program.cs` runs only the older samples in `Samples\DropMasters.cs` and `Samples\SetCustomProperties.cs`.

## Keeping the docs and the samples in step

```powershell
.\Check-DocsSamples.ps1 -DocsPath C:\path\to\VisioAutomation_GitBook_Docs
```

The script lists docs blocks with no sample, blocks whose text differs from their sample, and samples whose block is gone. It needs a checkout of the docs, which is why it is not part of the build.

When a docs block changes, change the matching sample in the same step, then build. When a block is added, removed or moved, the numbering after it shifts, so renumber the samples for that page (and the entries in `exempt.tsv`).

Docs for unreleased changes live on the docs repository's `next` branch (see `docs/decisions/gitbook-next-variant.md`). Check against the branch the samples were written for.

## Blocks that are not compiled

`Samples\Docs\exempt.tsv` lists each docs block that has no sample, with the reason: id, docs page, line, reason. Today these are API outlines (declarations shown for reference). An exempt block is not checked by the compiler, so keep the list short.

## Limits

- Compiling is not running. A sample can compile and still draw the wrong thing, and several assume setup the docs only describe.
- Compiling catches a renamed or removed member and a missing `using`. It does not catch a wrong explanation.
- Members marked `[Obsolete]` still compile, with a warning. Build warnings in this project mean the docs still show a name that is going away.
