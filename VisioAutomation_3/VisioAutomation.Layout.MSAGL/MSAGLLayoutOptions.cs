using VA=VisioAutomation;

namespace VisioAutomation.Layout.MSAGL
{
    public class MSAGLLayoutOptions : VA.Layout.DirectedGraph.LayoutOptions
    {
        public double ScalingFactor { get; set; }
        public bool UseDynamicConnectors { get; set; }

        public MSAGLLayoutOptions() :
            base()
        {
            UseDynamicConnectors = true;
            ScalingFactor = 14;
        }
    }
}