using IVisio = Microsoft.Office.Interop.Visio;

namespace VSamples.Docs.Samples.Docs
{
    // Code blocks from the .NET docs page text/paragraph.md. Compile-only: each method is one block,
    // with the names the block assumes (page, shape, client, ...) as parameters.
    public static class TextParagraphSamples
    {
        public static void Sample_01(IVisio.Shape shape)
        {
            // <sample id="text/paragraph#01" doc="text/paragraph.md" line="39">
            var pg = new VisioAutomation.Text.ParagraphCells();
            pg.HorizontalAlign = 1;       // center
            pg.SpacingBefore   = "6 pt";
            pg.SpacingAfter    = "6 pt";

            var writer = new VisioAutomation.ShapeSheet.Writers.SrcWriter();
            writer.SetValues(pg);
            writer.Commit(shape, VisioAutomation.Core.CellValueType.Formula);
            // </sample>
        }

        public static void Sample_02(IVisio.Shape shape)
        {
            // <sample id="text/paragraph#02" doc="text/paragraph.md" line="54">
            var rows = VisioAutomation.Text.ParagraphCells.GetCells(
                shape,
                VisioAutomation.Core.CellValueType.Result);

            foreach (var row in rows)
            {
                System.Console.WriteLine($"Align = {row.HorizontalAlign.Value}");
            }
            // </sample>
        }
    }
}
