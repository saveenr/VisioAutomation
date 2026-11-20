using IVisio = Microsoft.Office.Interop.Visio;

namespace VSamples.Docs.Samples.Docs
{
    // Code blocks from the .NET docs page visio-scripting/hyperlink.md. Compile-only: each method is one block,
    // with the names the block assumes (page, shape, client, ...) as parameters.
    public static class VisioScriptingHyperlinkSamples
    {
        public static void Sample_01(VisioScripting.Client client)
        {
            // <sample id="visio-scripting/hyperlink#01" doc="visio-scripting/hyperlink.md" line="17">
            var hlink = new VisioAutomation.Shapes.HyperlinkCells();
            hlink.Address = "\"https://example.com\"";
            hlink.Description = "\"Example\"";

            // Operate on whatever is currently selected
            var added = client.Hyperlink.AddHyperlink(VisioScripting.TargetShapes.Auto, hlink);

            var dict = client.Hyperlink.GetHyperlinks(VisioScripting.TargetShapes.Auto,
                                                       VisioAutomation.Core.CellValueType.Formula);
            foreach (var kv in dict)
            {
                foreach (var record in kv.Value)
                {
                    System.Console.WriteLine("{0}: {1}", kv.Key.Name, record.Address);
                }
            }

            client.Hyperlink.DeleteHyperlinkAtIndex(VisioScripting.TargetShapes.Auto, 0);
            // </sample>
        }
    }
}
