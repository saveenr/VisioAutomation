using System.Collections.Generic;

namespace VisioAutomation.Models.Layouts.Box
{
    public abstract class Node
    {
        internal Node parent;

        public object Data { get; set; }
        public Drawing.Rectangle Rectangle { get; set; }
        public Drawing.Rectangle ReservedRectangle { get; set; }
        public Drawing.Size Size { get; set; }
        public AlignmentHorizontal HAlignToParent;
        public AlignmentVertical VAlignToParent;
        public abstract Drawing.Size CalculateSize();
        public abstract void _place(Drawing.Point origin);
        public abstract IEnumerable<Node> GetChildren();

        public Node Parent
        {
            get { return this.parent; }
        }
    }
}