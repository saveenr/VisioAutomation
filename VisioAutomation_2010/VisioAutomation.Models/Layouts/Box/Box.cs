using System.Collections.Generic;

namespace VisioAutomation.Models.Layouts.Box
{
    public class Box : Node
    {
        public Box(double w, double h) :
            this(new VisioAutomation.Core.Size(w, h) )
        {
        }

        protected Box(VisioAutomation.Core.Size s)
        {
            this.Size = s;
        }

        public override VisioAutomation.Core.Size CalculateSize()
        {
            return this.Size;
        }

        public override void _place(VisioAutomation.Core.Point origin)
        {
            this.Rectangle = new VisioAutomation.Core.Rectangle(origin, this.Size);
        }

        public override IEnumerable<Node> GetChildren()
        {
            yield break;
        }
    }
}