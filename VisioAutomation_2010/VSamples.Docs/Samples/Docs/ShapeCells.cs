using IVisio = Microsoft.Office.Interop.Visio;

namespace VSamples.Docs.Samples.Docs
{
    // Code blocks from the .NET docs page shape-cells.md. Compile-only: each method is one block,
    // with the names the block assumes (page, shape, client, ...) as parameters.
    public static class ShapeCellsSamples
    {
        public static void Sample_01(IVisio.Shape shape, VisioAutomation.Shapes.ShapeXFormCells xform, VisioAutomation.Shapes.ShapeFormatCells fmt, VisioAutomation.Shapes.ShapeLayoutCells layout)
        {
            // <sample id="shape-cells#01" doc="shape-cells.md" line="17">
            var writer = new VisioAutomation.ShapeSheet.Writers.SrcWriter();
            writer.SetValues(xform);
            writer.SetValues(fmt);
            writer.SetValues(layout);
            writer.Commit(shape, VisioAutomation.Core.CellValueType.Formula);
            // </sample>
        }
    }
}
