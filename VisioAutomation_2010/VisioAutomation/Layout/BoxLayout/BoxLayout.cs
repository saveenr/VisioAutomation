using System.Collections.Generic;
using VA=VisioAutomation;

namespace VisioAutomation.Layout.BoxLayout
{
    public class BoxLayout
    {
        private Container _root;

        public BoxLayout()         
{
        }

        public Container Root
        {
            get { return _root; }
            set { _root = value; }
        }

        public void PerformLayout()
        {
            if (this.Root.Count < 1)
            {
                throw new VA.AutomationException("Root must contain at least one child");
            }

            this._root.CalculateSize();
            this.Place(new VA.Drawing.Point(0,0));
            this._root.ReservedRectangle = this._root.Rectangle;
        }

        private void Place(VA.Drawing.Point origin)
        {
            this._root._place(origin);
        }

        public IEnumerable<Node> Nodes
        {
            get
            {
                Node rootn = this._root;
                return VA.Internal.TreeTraversal.PreOrder(rootn, n => n.GetChildren());
            }
        }

    }
}