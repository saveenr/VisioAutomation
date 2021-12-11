using VisioAutomation.Models.Layouts.DirectedGraph;

namespace VSamples.Samples.Layouts
{
    public  class DirectedGraphViaMsagl : SampleMethodBase
    {
        public override void RunSample ()
        {
            var page1 = SampleEnvironment.Application.ActiveDocument.Pages.Add();
            var directed_graph_drawing = DirectedGraphViaVisio.get_dg_drawing();

            var renderer = new MsaglRenderer();
            renderer.LayoutOptions.UseDynamicConnectors = false;
            renderer.Render(page1, directed_graph_drawing);
        }
    }
}