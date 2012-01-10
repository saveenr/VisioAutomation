using System.Collections.Generic;
using VA = VisioAutomation;

namespace VisioAutomation.Layout.BoxLayout
{
    public abstract class Node
    {
        internal Node parent;
        public object Data { get; set; }
        public VA.Drawing.Rectangle Rectangle { get; set; }
        public VA.Drawing.Rectangle ReservedRectangle { get; set; }
        public VA.Drawing.Size Size { get; set; }
        public VA.Layout.BoxLayout.AlignmentHorizontal HAlignToParent;
        public VA.Layout.BoxLayout.AlignmentVertical VAlignToParent;

        public Node Parent
        {
            get { return this.parent; }
        }

        public abstract VA.Drawing.Size CalculateSize();
        public abstract void _place(VA.Drawing.Point origin);
        public abstract IEnumerable<Node> GetChildren();
    }
}