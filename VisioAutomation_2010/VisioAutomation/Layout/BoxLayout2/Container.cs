using System.Collections.Generic;
using VisioAutomation.Drawing;
using VA = VisioAutomation;
using System.Linq;

namespace VisioAutomation.Layout.BoxLayout2
{
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

        public Container AddColumnContainer(DirectionVertical vdir)
        {
            var n = new Container(ContainerDirection.Vertical);
            n.ChildVerticalDirection = vdir;
            this.AddNode(n);
            return n;
        }

        public Container AddRowContainer(DirectionHorizontal hdir)
        {
            var n = new Container(ContainerDirection.Horizontal);
            n.ChildHorizontalDirection = hdir;
            this.AddNode(n);
            return n;
        }

        public void AddNode(Node n)
        {
            if (n.Parent != null)
            {
                throw new VA.AutomationException("This item already has a parent");                
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

            double x = this.ChildHorizontalDirection == DirectionHorizontal.LeftToRight ? 
                origin.X + this.Padding 
                : origin.Y + this.Size.Width -Padding;

            double y = this.ChildVerticalDirection == DirectionVertical.BottomToTop ? 
                origin.Y + this.Padding 
                : origin.Y + this.Size.Height - this.Padding;

            double reserved_width = this.Size.Width - (2 * this.Padding);
            double reserved_height = this.Size.Height - (2 * this.Padding);
            foreach (var c in this.Children)
            {

                if (this.Direction == ContainerDirection.Vertical)
                {
                    double excess_width = reserved_width - c.Size.Width;
                    double align_delta_x = 0.0;

                    // If there is any excess width then we need to adjust
                    // for anyalignment
                    if (excess_width>0)
                    {
                        
                        if (c.HAlignToParent == AlignmentHorizontal.Left)
                        {
                            align_delta_x = 0;
                        }
                        else if (c.HAlignToParent == AlignmentHorizontal.Right)
                        {
                            align_delta_x = excess_width;
                        }
                        else if (c.HAlignToParent == AlignmentHorizontal.Center)
                        {
                            align_delta_x = excess_width / 2;
                        }
                    }


                    if (this.ChildVerticalDirection == DirectionVertical.BottomToTop)
                    {
                        // BOTTOM TO TOP
                        c.ReservedRectangle = new VA.Drawing.Rectangle(x, y, x + reserved_width, y + c.Size.Height);

                        c._place(new VA.Drawing.Point(x+align_delta_x, y));
                        y += c.Size.Height;
                        y += this.ChildSeparation;

                    }
                    else
                    {
                        // TOP TO BOTTOM
                        c.ReservedRectangle = new VA.Drawing.Rectangle(x, y - c.Size.Height, x + reserved_width, y);

                        c._place(new VA.Drawing.Point(x+align_delta_x, y - c.Size.Height));
                        y -= c.Size.Height;
                        y -= this.ChildSeparation;

                    }
                }
                else
                {
                    double excess_height = reserved_height - c.Size.Height;
                    double align_delta_y = 0.0;
                    // If there is any excess height then we need to adjust
                    // for any alignment
                    if (excess_height > 0)
                    {
                        if (c.VAlignToParent == AlignmentVertical.Bottom)
                        {
                            align_delta_y = 0;
                        }
                        else if (c.VAlignToParent == AlignmentVertical.Top)
                        {
                            align_delta_y = excess_height;
                        }
                        else if (c.VAlignToParent == AlignmentVertical.Center)
                        {
                            align_delta_y = excess_height / 2;
                        }
                    }

                    if (this.ChildHorizontalDirection == DirectionHorizontal.LeftToRight)
                    {
                        // LEFT TO RIGHT
                        c.ReservedRectangle = new VA.Drawing.Rectangle(x, y, x + c.Size.Width, y + reserved_height);

                        c._place(new VA.Drawing.Point(x, y+align_delta_y));
                        x += c.Size.Width;
                        x += this.ChildSeparation;

                    }
                    else 
                    {
                        // RIGHT TO LEFT
                        c.ReservedRectangle = new VA.Drawing.Rectangle(x - c.Size.Width, y, x, y + reserved_height);

                        c._place(new VA.Drawing.Point(x - c.Size.Width, y+align_delta_y));
                        x -= c.Size.Width;
                        x -= this.ChildSeparation;

                    }

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