using VA=VisioAutomation;

namespace VisioAutomation.Models.DirectedGraph
{
    public class MsaglLayoutOptions : LayoutOptions
    {
        public double ScalingFactor { get; set; }
        public bool UseDynamicConnectors { get; set; }

        public MsaglLayoutOptions() :
            base()
        {
            this.UseDynamicConnectors = true;
            this.ScalingFactor = 14;
        }
    }
}