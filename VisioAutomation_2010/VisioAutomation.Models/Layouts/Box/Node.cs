using System.Collections.Generic;

namespace VisioAutomation.Models.Layouts.Box
{
    public abstract class Node
    {
        public object Data { get; set; }
        public Geometry.Rectangle Rectangle { get; set; }
        public Geometry.Rectangle ReservedRectangle { get; set; }
        public Geometry.Size Size { get; set; }
        public AlignmentHorizontal HAlignToParent;
        public AlignmentVertical VAlignToParent;
        public abstract Geometry.Size CalculateSize();
        public abstract void _place(Geometry.Point origin);
        public abstract IEnumerable<Node> GetChildren();
    }
}