using IVisio = Microsoft.Office.Interop.Visio;

namespace VSamples.Docs.Samples.Docs
{
    // Code blocks from the .NET docs page shape-xform-cells.md. Compile-only: each method is one block,
    // with the names the block assumes (page, shape, client, ...) as parameters.
    public static class ShapeXformCellsSamples
    {
        public static void Sample_01(IVisio.Shape shape)
        {
            // <sample id="shape-xform-cells#01" doc="shape-xform-cells.md" line="23">
            var xform = new VisioAutomation.Shapes.ShapeXFormCells();
            xform.PinX   = 4.0;
            xform.PinY   = 2.0;
            xform.Width  = 1.5;
            xform.Height = 1.0;

            var writer = new VisioAutomation.ShapeSheet.Writers.SrcWriter();
            writer.SetValues(xform);
            writer.Commit(shape, VisioAutomation.Core.CellValueType.Formula);
            // </sample>
        }

        public static void Sample_02(IVisio.Shape shape)
        {
            // <sample id="shape-xform-cells#02" doc="shape-xform-cells.md" line="41">
            var xform = VisioAutomation.Shapes.ShapeXFormCells.GetCells(
                shape,
                VisioAutomation.Core.CellValueType.Result);

            System.Console.WriteLine($"Width = {xform.Width.Value}");
            // </sample>
        }
    }
}
