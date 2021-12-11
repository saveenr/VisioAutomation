using VisioAutomation.Models.Layouts.DirectedGraph;

namespace VSamples.Layouts
{
    public static class DirectedGraphViaMsaglX
    {
        public static void DirectedGraphViaMsagl()
        {
            var page1 = SampleEnvironment.Application.ActiveDocument.Pages.Add();
            var directed_graph_drawing = DirectedGraphViaVisioX.get_dg_drawing();

            var renderer = new MsaglRenderer();
            renderer.LayoutOptions.UseDynamicConnectors = false;
            renderer.Render(page1, directed_graph_drawing);
        }
    }
}