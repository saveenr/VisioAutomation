using VA=VisioAutomation;

namespace VisioAutomation.Models.DirectedGraph
{
    public class MsaglLayoutOptions : VA.Models.DirectedGraph.LayoutOptions
    {
        public double ScalingFactor { get; set; }
        public bool UseDynamicConnectors { get; set; }

        public MsaglLayoutOptions() :
            base()
        {
            UseDynamicConnectors = true;
            ScalingFactor = 14;
        }
    }
}