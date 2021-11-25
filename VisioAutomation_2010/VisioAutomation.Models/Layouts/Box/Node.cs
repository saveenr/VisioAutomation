using System.Collections.Generic;

namespace VisioAutomation.Models.Layouts.Box
{
    public abstract class Node
    {
        public object Data { get; set; }
        public VisioAutomation.Core.Rectangle Rectangle { get; set; }
        public VisioAutomation.Core.Rectangle ReservedRectangle { get; set; }
        public VisioAutomation.Core.Size Size { get; set; }
        public AlignmentHorizontal HAlignToParent;
        public AlignmentVertical VAlignToParent;
        public abstract VisioAutomation.Core.Size CalculateSize();
        public abstract void _place(VisioAutomation.Core.Point origin);
        public abstract IEnumerable<Node> GetChildren();
    }
}