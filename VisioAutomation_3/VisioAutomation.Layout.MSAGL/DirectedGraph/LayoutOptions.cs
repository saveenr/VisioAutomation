using VA=VisioAutomation;

namespace VisioAutomation.Layout.DirectedGraph
{
    public class LayoutOptions
    {
        public double ScalingFactor { get; set; }
        public VA.Drawing.Size ResizeBorderWidth { get; set; }
        public bool UseDynamicConnectors { get; set; }
        public bool HideConnectionPoints { get; set; }
        public bool HideGrid { get; set; }
        public VA.Drawing.Size DefaultShapeSize { get; set; }

        public LayoutOptions()
        {
            DefaultShapeSize = new VA.Drawing.Size(1.0, 0.75);
            HideGrid = true;
            HideConnectionPoints = true;
            UseDynamicConnectors = true;
            ResizeBorderWidth = new VA.Drawing.Size(0.5, 0.5);
            ScalingFactor = 14;
        }

    }
}