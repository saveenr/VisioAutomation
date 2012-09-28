using VA=VisioAutomation;

namespace VisioAutomation.Layout.Models.DirectedGraph
{
    public class LayoutOptions
    {
        public VA.Drawing.Size ResizeBorderWidth { get; set; }
        public bool HideConnectionPoints { get; set; }
        public bool HideGrid { get; set; }
        public VA.Drawing.Size DefaultShapeSize { get; set; }

        public LayoutOptions()
        {
            ResizeBorderWidth = new VA.Drawing.Size(0.5, 0.5);
            DefaultShapeSize = new VA.Drawing.Size(1.0, 0.75);
            HideConnectionPoints = true;
            HideGrid = true;            
        }        
    }
}
