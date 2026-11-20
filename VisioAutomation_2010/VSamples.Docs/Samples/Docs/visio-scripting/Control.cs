using IVisio = Microsoft.Office.Interop.Visio;

namespace VSamples.Docs.Samples.Docs
{
    // Code blocks from the .NET docs page visio-scripting/control.md. Compile-only: each method is one block,
    // with the names the block assumes (page, shape, client, ...) as parameters.
    public static class VisioScriptingControlSamples
    {
        public static void Sample_01(VisioScripting.Client client, IVisio.Shape shape, double a)
        {
            // <sample id="visio-scripting/control#01" doc="visio-scripting/control.md" line="17">
            var ctrl = new VisioAutomation.Shapes.ControlCells();
            ctrl.X = "Width*0.5";
            ctrl.Y = "Height*0.5";
            ctrl.XBehavior = "0";
            ctrl.YBehavior = "0";

            // Add a centered control handle to every selected shape
            var indices = client.Control.AddControlToShapes(VisioScripting.TargetShapes.Auto, ctrl);

            var dict = client.Control.GetControls(VisioScripting.TargetShapes.Auto,
                                                    VisioAutomation.Core.CellValueType.Formula);
            foreach (var kv in dict)
            {
                int row = 0;
                foreach (var record in kv.Value)
                {
                    System.Console.WriteLine("{0} ctl[{1}]: ({2}, {3})",
                        kv.Key.Name, row, record.X, record.Y);
                    row++;
                }
            }
            // </sample>
        }
    }
}
