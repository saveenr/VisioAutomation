using IVisio = Microsoft.Office.Interop.Visio;

namespace VSamples.Docs.Samples.Docs
{
    // Code blocks from the .NET docs page control-handles.md. Compile-only: each method is one block,
    // with the names the block assumes (page, shape, client, ...) as parameters.
    public static class ControlHandlesSamples
    {
        public static void Sample_01(IVisio.Shape shape, double a)
        {
            // <sample id="control-handles#01" doc="control-handles.md" line="21">
            // Default-positioned handle
            int row1 = VisioAutomation.Shapes.ControlHelper.Add(shape);

            // Centered handle with a tooltip
            var ctrl = new VisioAutomation.Shapes.ControlCells();
            ctrl.X = "Width*0.5";
            ctrl.Y = "Height*0.5";
            ctrl.Tip = "Adjust";
            int row2 = VisioAutomation.Shapes.ControlHelper.Add(shape, ctrl);
            // </sample>
        }

        public static void Sample_02(IVisio.Shape shape, int row)
        {
            // <sample id="control-handles#02" doc="control-handles.md" line="39">
            var ctrl = new VisioAutomation.Shapes.ControlCells();
            ctrl.Tip = "Resize me";

            VisioAutomation.Shapes.ControlHelper.Set(shape, (short)row, ctrl);
            // </sample>
        }

        public static void Sample_03(IVisio.Shape shape)
        {
            // <sample id="control-handles#03" doc="control-handles.md" line="50">
            var cells = VisioAutomation.Shapes.ControlCells.GetCells(
                shape,
                VisioAutomation.Core.CellValueType.Formula);
            // </sample>
        }

        public static void Sample_04(IVisio.Shape shape)
        {
            // <sample id="control-handles#04" doc="control-handles.md" line="58">
            int count = VisioAutomation.Shapes.ControlHelper.GetCount(shape);
            // </sample>
        }
    }
}
