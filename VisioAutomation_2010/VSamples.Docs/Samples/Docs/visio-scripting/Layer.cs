using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;

namespace VSamples.Docs.Samples.Docs
{
    // Code blocks from the .NET docs page visio-scripting/layer.md. Compile-only: each method is one block,
    // with the names the block assumes (page, shape, client, ...) as parameters.
    public static class VisioScriptingLayerSamples
    {
        public static void Sample_01(VisioScripting.Client client)
        {
            // <sample id="visio-scripting/layer#01" doc="visio-scripting/layer.md" line="24">
            var page = client.Page.GetActivePage();
            var layers = page.Layers.ToList();
            foreach (var layer in layers)
            {
                System.Console.WriteLine("Layer: {0}", layer.Name);
            }
            // </sample>
        }
    }
}
