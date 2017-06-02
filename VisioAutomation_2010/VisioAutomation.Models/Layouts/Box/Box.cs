using System.Collections.Generic;

namespace VisioAutomation.Models.Layouts.Box
{
    public class Box : Node
    {
        public Box(double w, double h) :
            this(new Geometry.Size(w, h) )
        {
        }

        protected Box(Geometry.Size s)
        {
            this.Size = s;
        }

        public override Geometry.Size CalculateSize()
        {
            return this.Size;
        }

        public override void _place(Geometry.Point origin)
        {
            this.Rectangle = new Geometry.Rectangle(origin, this.Size);
        }

        public override IEnumerable<Node> GetChildren()
        {
            yield break;
        }
    }
}