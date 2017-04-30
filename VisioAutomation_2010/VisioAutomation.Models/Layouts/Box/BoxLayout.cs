using System.Collections.Generic;

namespace VisioAutomation.Models.Layouts.Box
{
    public class BoxLayout
    {
        public Container Root { get; set; }

        public IEnumerable<Node> Nodes
        {
            get
            {
                Node rootn = this.Root;
                return GenTreeOps.Algorithms.PreOrder(rootn, n => n.GetChildren());
            }
        }

        public void PerformLayout()
        {
            if (this.Root.Count < 1)
            {
                throw new System.ArgumentException("Root must contain at least one child");
            }

            this.Root.CalculateSize();
            this.Place(new Geometry.Point(0, 0));
            this.Root.ReservedRectangle = this.Root.Rectangle;
        }

        private void Place(Geometry.Point origin)
        {
            this.Root._place(origin);
        }
    }
}