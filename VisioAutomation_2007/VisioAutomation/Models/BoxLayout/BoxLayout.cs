using System.Collections.Generic;
using VA = VisioAutomation;

namespace VisioAutomation.Models.BoxLayout
{
    public class BoxLayout
    {
        private Container _root;

        public Container Root
        {
            get { return _root; }
            set { _root = value; }
        }

        public IEnumerable<Node> Nodes
        {
            get
            {
                Node rootn = _root;
                return VA.Internal.TreeOps.PreOrder(rootn, n => n.GetChildren());
            }
        }

        public void PerformLayout()
        {
            if (Root.Count < 1)
            {
                throw new AutomationException("Root must contain at least one child");
            }

            _root.CalculateSize();
            Place(new VA.Drawing.Point(0, 0));
            _root.ReservedRectangle = _root.Rectangle;
        }

        private void Place(VA.Drawing.Point origin)
        {
            _root._place(origin);
        }
    }
}