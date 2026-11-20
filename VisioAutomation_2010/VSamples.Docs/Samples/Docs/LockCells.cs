using IVisio = Microsoft.Office.Interop.Visio;

namespace VSamples.Docs.Samples.Docs
{
    // Code blocks from the .NET docs page lock-cells.md. Compile-only: each method is one block,
    // with the names the block assumes (page, shape, client, ...) as parameters.
    public static class LockCellsSamples
    {
        public static void Sample_01(IVisio.Shape shape)
        {
            // <sample id="lock-cells#01" doc="lock-cells.md" line="38">
            var locks = new VisioAutomation.Shapes.LockCells();
            locks.MoveX  = 1;
            locks.MoveY  = 1;
            locks.Width  = 1;
            locks.Height = 1;

            var writer = new VisioAutomation.ShapeSheet.Writers.SrcWriter();
            writer.SetValues(locks);
            writer.Commit(shape, VisioAutomation.Core.CellValueType.Formula);
            // </sample>
        }

        public static void Sample_02(IVisio.Shape shape)
        {
            // <sample id="lock-cells#02" doc="lock-cells.md" line="56">
            var locks = VisioAutomation.Shapes.LockCells.GetCells(
                shape,
                VisioAutomation.Core.CellValueType.Formula);

            System.Console.WriteLine($"Width-locked: {locks.Width.Value}");
            // </sample>
        }

        public static void Sample_03(IVisio.Page page, System.Collections.Generic.IList<int> shapeids)
        {
            // <sample id="lock-cells#03" doc="lock-cells.md" line="66">
            var locks_per_shape = VisioAutomation.Shapes.LockCells.GetCells(
                page, shapeids, VisioAutomation.Core.CellValueType.Formula);
            // </sample>
        }
    }
}
