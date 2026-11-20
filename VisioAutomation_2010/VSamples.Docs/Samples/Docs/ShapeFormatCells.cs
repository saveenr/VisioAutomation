using IVisio = Microsoft.Office.Interop.Visio;

namespace VSamples.Docs.Samples.Docs
{
    // Code blocks from the .NET docs page shape-format-cells.md. Compile-only: each method is one block,
    // with the names the block assumes (page, shape, client, ...) as parameters.
    public static class ShapeFormatCellsSamples
    {
        public static void Sample_01(IVisio.Shape shape)
        {
            // <sample id="shape-format-cells#01" doc="shape-format-cells.md" line="53">
            var fmt = new VisioAutomation.Shapes.ShapeFormatCells();
            fmt.FillForeground = "rgb(255,128,0)";
            fmt.FillPattern    = 1;
            fmt.LineColor      = "rgb(0,0,0)";
            fmt.LineWeight     = "0.5 pt";
            fmt.LinePattern    = 1;

            var writer = new VisioAutomation.ShapeSheet.Writers.SrcWriter();
            writer.SetValues(fmt);
            writer.Commit(shape, VisioAutomation.Core.CellValueType.Formula);
            // </sample>
        }

        public static void Sample_02(IVisio.Shape shape)
        {
            // <sample id="shape-format-cells#02" doc="shape-format-cells.md" line="72">
            var fmt = VisioAutomation.Shapes.ShapeFormatCells.GetCells(
                shape,
                VisioAutomation.Core.CellValueType.Result);

            System.Console.WriteLine($"Fill foreground = {fmt.FillForeground.Value}");
            // </sample>
        }
    }
}
