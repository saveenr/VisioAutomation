using IVisio = Microsoft.Office.Interop.Visio;

namespace VSamples.Docs.Samples.Docs
{
    // Code blocks from the .NET docs page shape-layout-cells.md. Compile-only: each method is one block,
    // with the names the block assumes (page, shape, client, ...) as parameters.
    public static class ShapeLayoutCellsSamples
    {
        public static void Sample_01(IVisio.Shape shape)
        {
            // <sample id="shape-layout-cells#01" doc="shape-layout-cells.md" line="43">
            var layout = new VisioAutomation.Shapes.ShapeLayoutCells();
            layout.ShapePlaceStyle = 0;          // PlaceCenter
            layout.ShapePermeableX = 1;
            layout.ShapePermeableY = 1;

            var writer = new VisioAutomation.ShapeSheet.Writers.SrcWriter();
            writer.SetValues(layout);
            writer.Commit(shape, VisioAutomation.Core.CellValueType.Formula);
            // </sample>
        }

        public static void Sample_02(IVisio.Shape shape)
        {
            // <sample id="shape-layout-cells#02" doc="shape-layout-cells.md" line="58">
            var layout = VisioAutomation.Shapes.ShapeLayoutCells.GetCells(
                shape,
                VisioAutomation.Core.CellValueType.Formula);

            System.Console.WriteLine($"PlaceStyle = {layout.ShapePlaceStyle.Value}");
            // </sample>
        }
    }
}
