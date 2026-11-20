using IVisio = Microsoft.Office.Interop.Visio;

namespace VSamples.Docs.Samples.Docs
{
    // Code blocks from the .NET docs page visio-scripting/output.md. Compile-only: each method is one block,
    // with the names the block assumes (page, shape, client, ...) as parameters.
    public static class VisioScriptingOutputSamples
    {
        public static void Sample_01(VisioScripting.Client client, System.Collections.Generic.List<IVisio.Shape> shapes, int count)
        {
            // <sample id="visio-scripting/output#01" doc="visio-scripting/output.md" line="26">
            client.Output.WriteVerbose("About to draw {0} shapes", count);

            for (int i = 0; i < count; i++)
            {
                var rect = new VisioAutomation.Core.Rectangle(i, 0, i + 1, 1);
                client.Draw.DrawRectangle(VisioScripting.TargetPage.Auto, rect);
            }

            client.Output.WriteVerbose("Done");
            // </sample>
        }
    }
}
