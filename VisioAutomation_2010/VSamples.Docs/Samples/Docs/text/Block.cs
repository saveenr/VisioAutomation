using IVisio = Microsoft.Office.Interop.Visio;

namespace VSamples.Docs.Samples.Docs
{
    // Code blocks from the .NET docs page text/block.md. Compile-only: each method is one block,
    // with the names the block assumes (page, shape, client, ...) as parameters.
    public static class TextBlockSamples
    {
        public static void Sample_01(IVisio.Shape shape)
        {
            // <sample id="text/block#01" doc="text/block.md" line="27">
            var tb = new VisioAutomation.Text.TextBlockCells();
            tb.LeftMargin    = "4 pt";
            tb.RightMargin   = "4 pt";
            tb.TopMargin     = "2 pt";
            tb.BottomMargin  = "2 pt";
            tb.VerticalAlign = 1;          // middle

            var writer = new VisioAutomation.ShapeSheet.Writers.SrcWriter();
            writer.SetValues(tb);
            writer.Commit(shape, VisioAutomation.Core.CellValueType.Formula);
            // </sample>
        }

        public static void Sample_02(IVisio.Shape shape)
        {
            // <sample id="text/block#02" doc="text/block.md" line="44">
            var tb = VisioAutomation.Text.TextBlockCells.GetTextBlockCells(
                shape,
                VisioAutomation.Core.CellValueType.Result);

            System.Console.WriteLine($"VerticalAlign = {tb.VerticalAlign.Value}");
            // </sample>
        }

        public static void Sample_03(IVisio.Page page, System.Collections.Generic.List<IVisio.Shape> shapes, IVisio.Shape shape1, IVisio.Shape shape2, double a)
        {
            // <sample id="text/block#03" doc="text/block.md" line="52">
            // Many shapes on a page in one batched call:
            var shapeids = new[] { shape1.ID, shape2.ID };
            var rows = VisioAutomation.Text.TextBlockCells.GetTextBlockCells(
                page, shapeids, VisioAutomation.Core.CellValueType.Result);
            // </sample>
        }
    }
}
