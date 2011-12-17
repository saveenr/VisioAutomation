using VA=VisioAutomation;

namespace VisioAutomation.Layout.MSAGL
{
    public class LayoutOptions : VA.Layout.DirectedGraph.LayoutOptions
    {
        public double ScalingFactor { get; set; }
        public bool UseDynamicConnectors { get; set; }

        public LayoutOptions() :
            base()
        {
            UseDynamicConnectors = true;
            ScalingFactor = 14;
        }
    }
}