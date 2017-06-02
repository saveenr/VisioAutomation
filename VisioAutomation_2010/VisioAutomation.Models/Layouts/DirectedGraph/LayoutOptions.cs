using VA=VisioAutomation;

namespace VisioAutomation.Models.Layouts.DirectedGraph
{
    public class LayoutOptions
    {
        public VA.Geometry.Size ResizeBorderWidth { get; set; }
        public VA.Geometry.Size DefaultShapeSize { get; set; }
        public LayoutDirection LayoutDirection { get; set; }

        public LayoutOptions()
        {
            this.ResizeBorderWidth = new VA.Geometry.Size(0.5, 0.5);
            this.DefaultShapeSize = new VA.Geometry.Size(1.0, 0.75);
            this.LayoutDirection = LayoutDirection.TopToBottom;
        }        
    }
}
