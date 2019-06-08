using VA=VisioAutomation;

namespace VisioAutomation.Models.Layouts.DirectedGraph
{
    public class DirectedGraphLayoutOptions
    {
        public VA.Geometry.Size ResizeBorderWidth { get; set; }
        public VA.Geometry.Size DefaultShapeSize { get; set; }
        public DirectedGraphLayoutDirection LayoutDirection { get; set; }

        internal DirectedGraphLayoutOptions()
        {
            this.ResizeBorderWidth = new VA.Geometry.Size(0.5, 0.5);
            this.DefaultShapeSize = new VA.Geometry.Size(1.0, 0.75);
            this.LayoutDirection = DirectedGraphLayoutDirection.TopToBottom;
        }        
    }
}
