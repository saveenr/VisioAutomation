using System.Collections.Generic;
using VisioAutomation.Drawing;
using VA = VisioAutomation;
using System.Linq;

namespace VisioAutomation.Layout.BoxLayout2
{

    public class Box : Node
    {
        public Box(double w, double h) :
            this(new VA.Drawing.Size(w, h) )
        {
        }

        protected Box(VA.Drawing.Size s)
        {
            this.Size = s;
        }

        public override VA.Drawing.Size CalculateSize()
        {
            return this.Size;
        }

        public override void _place(VA.Drawing.Point origin)
        {
            this.Rectangle = new VA.Drawing.Rectangle(origin, this.Size);
        }

        public override IEnumerable<Node> GetChildren()
        {
            yield break;
        }
    }
}