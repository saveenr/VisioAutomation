using IVisio = Microsoft.Office.Interop.Visio;

namespace VSamples.Docs.Samples.Docs
{
    // Code blocks from the .NET docs page analyzers.md. Compile-only: each method is one block,
    // with the names the block assumes (page, shape, client, ...) as parameters.
    public static class AnalyzersSamples
    {
        public static void Sample_01(IVisio.Page page)
        {
            // <sample id="analyzers#01" doc="analyzers.md" line="9">
            var options = new VisioAutomation.Analyzers.ConnectionAnalyzerOptions();
            var edges = VisioAutomation.Analyzers.ConnectionAnalyzer.GetDirectedEdges(page, options);

            foreach (var edge in edges)
            {
                System.Console.WriteLine(
                    "{0} -> {1} via {2}",
                    edge.From.NameU,
                    edge.To.NameU,
                    edge.Connector.NameU);
            }
            // </sample>
        }

        public static void Sample_02(IVisio.Page page)
        {
            // <sample id="analyzers#02" doc="analyzers.md" line="41">
            var options = new VisioAutomation.Analyzers.ConnectionAnalyzerOptions();
            options.EdgeDirectionSource    = VisioAutomation.Analyzers.EdgeDirectionSource.UseConnectorArrows;
            options.EdgeNoArrowsHandling   = VisioAutomation.Analyzers.EdgeNoArrowsHandling.ExcludeEdge;

            var edges = VisioAutomation.Analyzers.ConnectionAnalyzer.GetDirectedEdges(page, options);
            // </sample>
        }
    }
}
