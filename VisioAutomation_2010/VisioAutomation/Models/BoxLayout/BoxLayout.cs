using System.Collections.Generic;
using VA = VisioAutomation;

namespace VisioAutomation.Models.BoxLayout
{
    public class BoxLayout
    {
        private Container _root;

        public Container Root
        {
            get { return this._root; }
            set { this._root = value; }
        }

        public IEnumerable<Node> Nodes
        {
            get
            {
                Node rootn = this._root;
                return Internal.TreeOps.PreOrder(rootn, n => n.GetChildren());
            }
        }

        public void PerformLayout()
        {
            if (this.Root.Count < 1)
            {
                throw new AutomationException("Root must contain at least one child");
            }

            this._root.CalculateSize();
            this.Place(new Drawing.Point(0, 0));
            this._root.ReservedRectangle = this._root.Rectangle;
        }

        private void Place(Drawing.Point origin)
        {
            this._root._place(origin);
        }
    }
}