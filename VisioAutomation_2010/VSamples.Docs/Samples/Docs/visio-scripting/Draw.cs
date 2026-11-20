using IVisio = Microsoft.Office.Interop.Visio;

namespace VSamples.Docs.Samples.Docs
{
    // Code blocks from the .NET docs page visio-scripting/draw.md. Compile-only: each method is one block,
    // with the names the block assumes (page, shape, client, ...) as parameters.
    public static class VisioScriptingDrawSamples
    {
        public static void Sample_01(VisioScripting.Client client, IVisio.Page page)
        {
            // <sample id="visio-scripting/draw#01" doc="visio-scripting/draw.md" line="30">
            // Rectangle, oval, line, polyline on the active page
            var rect = client.Draw.DrawRectangle(VisioScripting.TargetPage.Auto, 0, 0, 4, 2);
            var oval = client.Draw.DrawOval(VisioScripting.TargetPage.Auto, 5, 0, 9, 2);
            var line = client.Draw.DrawLine(VisioScripting.TargetPage.Auto, 0, 3, 9, 3);

            var points = new[]
            {
                new VisioAutomation.Core.Point(0, 4),
                new VisioAutomation.Core.Point(2, 5),
                new VisioAutomation.Core.Point(4, 4),
                new VisioAutomation.Core.Point(6, 5),
            };
            var poly = client.Draw.DrawPolyLine(VisioScripting.TargetPage.Auto, points);
            // </sample>
        }
    }
}
