using System.Collections.Generic;
using VA = VisioAutomation;

namespace VisioAutomation.Models.BoxLayout
{
    public abstract class Node
    {
        internal Node parent;

        public object Data { get; set; }
        public VA.Drawing.Rectangle Rectangle { get; set; }
        public VA.Drawing.Rectangle ReservedRectangle { get; set; }
        public VA.Drawing.Size Size { get; set; }
        public VA.Models.BoxLayout.AlignmentHorizontal HAlignToParent;
        public VA.Models.BoxLayout.AlignmentVertical VAlignToParent;
        public abstract VA.Drawing.Size CalculateSize();
        public abstract void _place(VA.Drawing.Point origin);
        public abstract IEnumerable<Node> GetChildren();

        public Node Parent
        {
            get { return this.parent; }
        }
    }
}