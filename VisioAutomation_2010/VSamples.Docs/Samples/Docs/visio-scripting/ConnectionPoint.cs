using IVisio = Microsoft.Office.Interop.Visio;

namespace VSamples.Docs.Samples.Docs
{
    // Code blocks from the .NET docs page visio-scripting/connection-point.md. Compile-only: each method is one block,
    // with the names the block assumes (page, shape, client, ...) as parameters.
    public static class VisioScriptingConnectionPointSamples
    {
        public static void Sample_01(VisioScripting.Client client, IVisio.Shape shape)
        {
            // <sample id="visio-scripting/connection-point#01" doc="visio-scripting/connection-point.md" line="17">
            // Two new points on every selected shape: midpoint of left edge, midpoint of right edge
            client.ConnectionPoint.AddConnectionPoint(VisioScripting.TargetShapes.Auto,
                                                        "Width*0",
                                                        "Height*0.5",
                                                        VisioScripting.Models.ConnectionPointType.Inward);

            client.ConnectionPoint.AddConnectionPoint(VisioScripting.TargetShapes.Auto,
                                                        "Width*1",
                                                        "Height*0.5",
                                                        VisioScripting.Models.ConnectionPointType.Inward);

            // Inspect them
            var dict = client.ConnectionPoint.GetConnectionPoints(VisioScripting.TargetShapes.Auto);
            foreach (var kv in dict)
            {
                int row = 0;
                foreach (var record in kv.Value)
                {
                    System.Console.WriteLine("{0} cp[{1}]: ({2}, {3})",
                        kv.Key.Name, row, record.X, record.Y);
                    row++;
                }
            }
            // </sample>
        }
    }
}
