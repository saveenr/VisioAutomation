using System.Collections.Generic;

namespace VisioAutomation.Models.Layouts.Box
{
    public class Box : Node
    {
        public Box(double w, double h) :
            this(new Drawing.Size(w, h) )
        {
        }

        protected Box(Drawing.Size s)
        {
            this.Size = s;
        }

        public override Drawing.Size CalculateSize()
        {
            return this.Size;
        }

        public override void _place(Drawing.Point origin)
        {
            this.Rectangle = new Drawing.Rectangle(origin, this.Size);
        }

        public override IEnumerable<Node> GetChildren()
        {
            yield break;
        }
    }
}