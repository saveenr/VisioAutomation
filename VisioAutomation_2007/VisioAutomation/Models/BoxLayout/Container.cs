using System.Collections.Generic;
using System.Collections;
using VA = VisioAutomation;

namespace VisioAutomation.Models.BoxLayout
{
    public class Container : Node, IEnumerable<Node>
    {
        private List<Node> m_children;
        public double PaddingTop { get; set; }
        public double PaddingLeft { get; set; }
        public double PaddingRight{ get; set; }
        public double PaddingBottom { get; set; }
        public double ChildSpacing { get; set; }
        public Direction Direction { get; set; }
        public double MinWidth { get; set; }
        public double MinHeight { get; set; }

        public Container(Direction dir)
            : this(dir, 0,0)
        {
        }

        public Container(Direction dir, double minwidth, double minheight)
        {
            this.Direction = dir;
            this.PaddingLeft = 0.125;
            this.PaddingRight = 0.125;
            this.PaddingTop = 0.125;
            this.PaddingBottom = 0.125;
            this.ChildSpacing = 0.125;
            this.MinWidth = minwidth;
            this.MinHeight = minheight;
        }

        public IEnumerator<Node> GetEnumerator()
        {
            if (this.m_children == null)
            {
                yield break;
            }

            foreach (var c in this.m_children)
            {
                yield return c;
            }
        }

        IEnumerator IEnumerable.GetEnumerator()     
        {                                           
            return GetEnumerator();
        }

        public Box AddBox(double w, double h)
        {
            var n = new Box(w, h);
            this.AddNode(n);
            return n;
        }

        public Container AddContainer(Direction dir)
        {
            return this.AddContainer(dir,0,0);
        }

        public Container AddContainer(Direction dir, double minwidth, double minheight)
        {
            var n = new Container(dir,minwidth,minheight);
            n.Direction = dir;
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

        private bool is_hor()
        {
            return (this.Direction == Direction.LeftToRight) || (this.Direction == Direction.RightToLeft);
        }

        private bool is_ver()
        {
            return (this.Direction == Direction.TopToBottom) || (this.Direction == Direction.BottomToTop);
        }

        public override VA.Drawing.Size  CalculateSize()
        {
            double w = this.MinWidth;
            double h = this.MinHeight;

            double max_child_width = 0;
            double max_child_height = 0;
            double total_child_width = 0;
            double total_child_height = 0;

            foreach (var c in this)
            {
                var s = c.CalculateSize();
                max_child_width = System.Math.Max(max_child_width , s.Width);
                max_child_height = System.Math.Max(max_child_height, s.Height);
                total_child_height += s.Height;
                total_child_width += s.Width;
            }

            if ( this.is_hor())
            {
                w = System.Math.Max(w, total_child_width);
                h = System.Math.Max(h, max_child_height);
            }
            else
            {
                w = System.Math.Max(w, max_child_width);
                h = System.Math.Max(h, total_child_height);
            }
            
            w += this.PaddingLeft + this.PaddingRight;
            h += this.PaddingTop + this.PaddingBottom;

            // Account for child separation
            int num_seps = System.Math.Max(0, this.Count - 1);
            double total_sepy = (this.is_ver()) ? num_seps * this.ChildSpacing : 0.0;
            double total_sepx = (this.is_hor()) ? num_seps * this.ChildSpacing : 0.0;

            w += total_sepx;
            h += total_sepy;

            this.Size = new VA.Drawing.Size(w, h);
            return this.Size;
        }

        public override void _place(VA.Drawing.Point origin)
        {
            this.Rectangle = new VA.Drawing.Rectangle(origin, this.Size);

            double x;
            double y;

            if (this.Direction == Direction.RightToLeft)
            {
                x = origin.Y + this.Size.Width - this.PaddingRight;
            }
            else
            {
                x = origin.X + this.PaddingLeft;                
            }


            if (this.Direction == Direction.TopToBottom)
            {
                y = origin.Y + this.Size.Height - this.PaddingTop;
            }
            else
            {
                y = origin.Y + this.PaddingBottom;
            }

            double reserved_width = this.Size.Width - (this.PaddingLeft + this.PaddingRight);
            double reserved_height = this.Size.Height - (this.PaddingTop + this.PaddingBottom);
            foreach (var c in this)
            {

                if (this.is_ver())
                {
                    double excess_width = reserved_width - c.Size.Width;
                    double align_delta_x = 0.0;

                    // If there is any excess width then we need to adjust
                    // for anyalignment
                    if (excess_width>0)
                    {

                        if (c.HAlignToParent == VA.Models.BoxLayout.AlignmentHorizontal.Left)
                        {
                            align_delta_x = 0;
                        }
                        else if (c.HAlignToParent == VA.Models.BoxLayout.AlignmentHorizontal.Right)
                        {
                            align_delta_x = excess_width;
                        }
                        else if (c.HAlignToParent == VA.Models.BoxLayout.AlignmentHorizontal.Center)
                        {
                            align_delta_x = excess_width / 2;
                        }
                    }


                    if (this.Direction == Direction.BottomToTop)
                    {
                        // BOTTOM TO TOP
                        c.ReservedRectangle = new VA.Drawing.Rectangle(x, y, x + reserved_width, y + c.Size.Height);

                        c._place(new VA.Drawing.Point(x+align_delta_x, y));
                        y += c.Size.Height;
                        y += this.ChildSpacing;

                    }
                    else
                    {
                        // TOP TO BOTTOM
                        c.ReservedRectangle = new VA.Drawing.Rectangle(x, y - c.Size.Height, x + reserved_width, y);

                        c._place(new VA.Drawing.Point(x+align_delta_x, y - c.Size.Height));
                        y -= c.Size.Height;
                        y -= this.ChildSpacing;

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
                        if (c.VAlignToParent == VA.Models.BoxLayout.AlignmentVertical.Bottom)
                        {
                            align_delta_y = 0;
                        }
                        else if (c.VAlignToParent == VA.Models.BoxLayout.AlignmentVertical.Top)
                        {
                            align_delta_y = excess_height;
                        }
                        else if (c.VAlignToParent == VA.Models.BoxLayout.AlignmentVertical.Center)
                        {
                            align_delta_y = excess_height / 2;
                        }
                    }

                    if (this.Direction == Direction.LeftToRight)
                    {
                        // LEFT TO RIGHT
                        c.ReservedRectangle = new VA.Drawing.Rectangle(x, y, x + c.Size.Width, y + reserved_height);

                        c._place(new VA.Drawing.Point(x, y+align_delta_y));
                        x += c.Size.Width;
                        x += this.ChildSpacing;

                    }
                    else 
                    {
                        // RIGHT TO LEFT
                        c.ReservedRectangle = new VA.Drawing.Rectangle(x - c.Size.Width, y, x, y + reserved_height);

                        c._place(new VA.Drawing.Point(x - c.Size.Width, y+align_delta_y));
                        x -= c.Size.Width;
                        x -= this.ChildSpacing;

                    }

                }
            }
        }

        public override IEnumerable<Node> GetChildren()
        {
            foreach (var c in this)
            {
                yield return c;
            }
        }
    }
}