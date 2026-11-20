using IVisio = Microsoft.Office.Interop.Visio;

namespace VSamples.Docs.Samples.Docs
{
    // Code blocks from the .NET docs page connection-points.md. Compile-only: each method is one block,
    // with the names the block assumes (page, shape, client, ...) as parameters.
    public static class ConnectionPointsSamples
    {
        public static void Sample_01(IVisio.Shape shape)
        {
            // <sample id="connection-points#01" doc="connection-points.md" line="19">
            var cp = new VisioAutomation.Shapes.ConnectionPointCells();
            cp.X = "Width*0.5";
            cp.Y = "Height";

            int row = VisioAutomation.Shapes.ConnectionPointHelper.Add(shape, cp);
            // </sample>
        }

        public static void Sample_02(IVisio.Shape shape, int row)
        {
            // <sample id="connection-points#02" doc="connection-points.md" line="31">
            var cp = new VisioAutomation.Shapes.ConnectionPointCells();
            cp.X = "Width";
            cp.Y = "Height*0.5";

            VisioAutomation.Shapes.ConnectionPointHelper.Set(shape, (short)row, cp);
            // </sample>
        }

        public static void Sample_03(IVisio.Shape shape)
        {
            // <sample id="connection-points#03" doc="connection-points.md" line="43">
            var cp = VisioAutomation.Shapes.ConnectionPointCells.GetCells(
                shape,
                VisioAutomation.Core.CellValueType.Formula);
            // </sample>
        }

        public static void Sample_04(IVisio.Shape shape)
        {
            // <sample id="connection-points#04" doc="connection-points.md" line="51">
            int count = VisioAutomation.Shapes.ConnectionPointHelper.GetCount(shape);
            // </sample>
        }
    }
}
