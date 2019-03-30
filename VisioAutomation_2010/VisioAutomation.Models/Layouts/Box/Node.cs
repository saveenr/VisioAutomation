using System.Collections.Generic;

namespace VisioAutomation.Models.Layouts.Box
{
    public abstract class Node
    {
        public object Data { get; set; }
        public VisioAutomation.Geometry.Rectangle Rectangle { get; set; }
        public VisioAutomation.Geometry.Rectangle ReservedRectangle { get; set; }
        public VisioAutomation.Geometry.Size Size { get; set; }
        public AlignmentHorizontal HAlignToParent;
        public AlignmentVertical VAlignToParent;
        public abstract VisioAutomation.Geometry.Size CalculateSize();
        public abstract void _place(VisioAutomation.Geometry.Point origin);
        public abstract IEnumerable<Node> GetChildren();
    }
}