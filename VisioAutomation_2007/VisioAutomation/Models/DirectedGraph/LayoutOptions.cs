using VA=VisioAutomation;

namespace VisioAutomation.Models.DirectedGraph
{
    public class LayoutOptions
    {
        public VA.Drawing.Size ResizeBorderWidth { get; set; }
        public VA.Drawing.Size DefaultShapeSize { get; set; }
        public LayoutDirection LayoutDirection { get; set; }

        public LayoutOptions()
        {
            ResizeBorderWidth = new VA.Drawing.Size(0.5, 0.5);
            DefaultShapeSize = new VA.Drawing.Size(1.0, 0.75);
            this.LayoutDirection = LayoutDirection.TopToBottom;
        }        
    }
}
