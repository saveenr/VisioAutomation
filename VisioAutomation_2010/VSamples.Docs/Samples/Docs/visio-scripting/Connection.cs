using IVisio = Microsoft.Office.Interop.Visio;

namespace VSamples.Docs.Samples.Docs
{
    // Code blocks from the .NET docs page visio-scripting/connection.md. Compile-only: each method is one block,
    // with the names the block assumes (page, shape, client, ...) as parameters.
    public static class VisioScriptingConnectionSamples
    {
        public static void Sample_01(VisioScripting.Client client, IVisio.Page page, System.Collections.Generic.List<IVisio.Shape> shapes)
        {
            // <sample id="visio-scripting/connection#01" doc="visio-scripting/connection.md" line="24">
            // Three shapes already on the active page: a, b, c
            var a = client.Draw.DrawRectangle(VisioScripting.TargetPage.Auto, 0, 0, 1, 1);
            var b = client.Draw.DrawRectangle(VisioScripting.TargetPage.Auto, 3, 0, 4, 1);
            var c = client.Draw.DrawRectangle(VisioScripting.TargetPage.Auto, 6, 0, 7, 1);

            // Connect a -> b and b -> c with the default connector
            client.Connection.ConnectShapes(VisioScripting.TargetPage.Auto,
                                              new[] { a, b },
                                              new[] { b, c },
                                              null);

            // Inspect the resulting directed-edge graph
            var edges = client.Connection.GetDirectedEdgesOnPage(
                VisioScripting.TargetPage.Auto,
                new VisioAutomation.Analyzers.ConnectionAnalyzerOptions());

            foreach (var e in edges)
            {
                System.Console.WriteLine("{0} -> {1}", e.From.Name, e.To.Name);
            }
            // </sample>
        }
    }
}
