using IVisio = Microsoft.Office.Interop.Visio;

namespace VSamples.Docs.Samples.Docs
{
    // Code blocks from the .NET docs page text-formatting.md. Compile-only: each method is one block,
    // with the names the block assumes (page, shape, client, ...) as parameters.
    public static class TextFormattingSamples
    {
        public static void Sample_01(IVisio.Shape shape, VisioAutomation.Text.CharacterCells ch, VisioAutomation.Text.ParagraphCells pg, VisioAutomation.Text.TextBlockCells tb)
        {
            // <sample id="text-formatting#01" doc="text-formatting.md" line="27">
            var writer = new VisioAutomation.ShapeSheet.Writers.SrcWriter();
            writer.SetValues(ch);
            writer.SetValues(pg);
            writer.SetValues(tb);
            writer.Commit(shape, VisioAutomation.Core.CellValueType.Formula);
            // </sample>
        }
    }
}
