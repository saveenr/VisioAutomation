using System.Collections.Generic;
using VA = VisioAutomation;
using System.Linq;

namespace VisioAutomation.Layout.BoxLayout2
{

    public abstract class Node
    {
        internal Node parent;
        public object Data { get; set; }
        public VA.Drawing.Rectangle Rectangle { get; set; }
        public VA.Drawing.Size Size { get; set; }
        public VA.Drawing.AlignmentHorizontal HAlignToParent;
        public VA.Drawing.AlignmentVertical VAlignToParent;

        public Node Parent
        {
            get { return this.parent; }
        }

        public abstract VA.Drawing.Size CalculateSize();
        public abstract void _place(VA.Drawing.Point origin);
        public abstract IEnumerable<Node> GetChildren();
    }

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


    public class Container : Node
    {
        private List<Node> m_children;
        public double Padding { get; set; }
        public double ChildSeparation { get; set; }
        public ContainerDirection Direction;
        public DirectionVertical ChildVerticalDirection;
        public DirectionHorizontal ChildHorizontalDirection;
        public double MinWidth;
        public double MinHeight;

        public Container(ContainerDirection dir)
        {
            this.Direction = dir;
            this.ChildVerticalDirection = DirectionVertical.BottomToTop;
            this.ChildHorizontalDirection = DirectionHorizontal.LeftToRight;
        }

        public IEnumerable<Node> Children
        {
            get
            {
                if (this.m_children == null)
                {
                    yield break;
                }
                else
                {
                    foreach (var c in this.m_children)
                    {
                        yield return c;
                    }
                }
            }
        }

        public Box AddBox(double w, double h)
        {
            var n = new Box(w, h);
            this.AddNode(n);
            return n;
        }

        public void AddNode(Node n)
        {
            if (n.Parent != null)
            {
                throw new VA.AutomationException("This item has not been positioned");                
            }

            if (this.m_children == null)
            {
                this.m_children = new List<Node>();
            }

            this.m_children.Add(n);
        }

        public int Count
        {
            get
            {
                if (this.m_children == null)
                {
                    return 0;
                }
                else
                {
                    return this.m_children.Count;
                }
            }
        }

        public override VA.Drawing.Size  CalculateSize()
        {
            double w = this.MinWidth;
            double h = this.MinHeight;

            double max_child_width = 0;
            double max_child_height = 0;
            double total_child_width = 0;
            double total_child_height = 0;

            foreach (var c in this.Children)
            {
                var s = c.CalculateSize();
                max_child_width = System.Math.Max(max_child_width , s.Width);
                max_child_height = System.Math.Max(max_child_height, s.Height);
                total_child_height += s.Height;
                total_child_width += s.Width;
            }

            if (Direction == ContainerDirection.Horizontal)
            {
                w = System.Math.Max(w, total_child_width);
                h = System.Math.Max(h, max_child_height);
            }
            else
            {
                w = System.Math.Max(w, max_child_width);
                h = System.Math.Max(h, total_child_height);
            }
            
            w += (2 * this.Padding);
            h += (2 * this.Padding);

            // Account for child separation
            int num_seps = System.Math.Max(0, this.Count - 1);
            double total_sepy = (this.Direction == ContainerDirection.Vertical) ? num_seps * this.ChildSeparation : 0.0;
            double total_sepx = (this.Direction == ContainerDirection.Horizontal) ? num_seps * this.ChildSeparation : 0.0;

            w += total_sepx;
            h += total_sepy;

            this.Size = new VA.Drawing.Size(w, h);
            return this.Size;
        }

        public override void _place(VA.Drawing.Point origin)
        {
            this.Rectangle = new VA.Drawing.Rectangle(origin, this.Size);

            double x = origin.X + this.Padding;
            double y = origin.Y + this.Padding;
            foreach (var c in this.Children)
            {
                c._place( new VA.Drawing.Point(x,y));

                if (this.Direction == ContainerDirection.Vertical)
                {
                    y += c.Size.Height;
                    y += this.ChildSeparation;
                }
                else
                {
                    x += c.Size.Width;
                    x += this.ChildSeparation;
                }
            }
        }

        public override IEnumerable<Node> GetChildren()
        {
            foreach (var c in this.Children)
            {
                yield return c;
            }
        }

    }

}