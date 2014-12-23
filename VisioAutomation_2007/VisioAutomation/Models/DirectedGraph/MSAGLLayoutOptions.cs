using VA=VisioAutomation;

namespace VisioAutomation.Models.DirectedGraph
{
    public class MSAGLLayoutOptions : VA.Models.DirectedGraph.LayoutOptions
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