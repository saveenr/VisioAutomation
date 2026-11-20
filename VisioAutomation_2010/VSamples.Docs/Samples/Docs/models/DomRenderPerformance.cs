using IVisio = Microsoft.Office.Interop.Visio;
using VADOM = VisioAutomation.Models.Dom;

namespace VSamples.Docs.Samples.Docs
{
    // Code blocks from the .NET docs page models/dom-render-performance.md. Compile-only: each method is one block,
    // with the names the block assumes (page, shape, client, ...) as parameters.
    public static class ModelsDomRenderPerformanceSamples
    {
        public static void Sample_01()
        {
            // <sample id="models/dom-render-performance#01" doc="models/dom-render-performance.md" line="44">
            var page_node = new VADOM.Page();
            page_node.RenderPerformanceSettings.DeferRecalc = 1;   // adjust before Render
            // </sample>
        }
    }
}
