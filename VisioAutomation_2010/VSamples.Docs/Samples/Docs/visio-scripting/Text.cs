using IVisio = Microsoft.Office.Interop.Visio;

namespace VSamples.Docs.Samples.Docs
{
    // Code blocks from the .NET docs page visio-scripting/text.md. Compile-only: each method is one block,
    // with the names the block assumes (page, shape, client, ...) as parameters.
    public static class VisioScriptingTextSamples
    {
        public static void Sample_01(VisioScripting.Client client, IVisio.Shape shape, System.Collections.Generic.List<IVisio.Shape> shapes)
        {
            // <sample id="visio-scripting/text#01" doc="visio-scripting/text.md" line="24">
            // Set distinct labels on three selected shapes
            client.Text.SetShapeText(VisioScripting.TargetShapes.Auto,
                                       new[] { "Start", "Middle", "End" });

            // Set one label across all selected shapes (cyclic reuse)
            client.Text.SetShapeText(VisioScripting.TargetShapes.Auto,
                                       new[] { "Step" });

            // Read every selected shape's text back
            var texts = client.Text.GetShapeText(VisioScripting.TargetShapes.Auto);
            foreach (var t in texts)
            {
                System.Console.WriteLine(t);
            }
            // </sample>
        }
    }
}
