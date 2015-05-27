using System.Collections.Generic;
using VA=VisioAutomation;

namespace VisioAutomation.Drawing
{
    public struct BoundingBox
    {
        private bool initialized;

        private double min_x;
        private double min_y;
        private double max_x;
        private double max_y;

        public BoundingBox( IEnumerable<Point> points) :
            this()
        {
            foreach (var p in points)
            {
                this.Add(p);
            }
        }

        public BoundingBox(IEnumerable<Rectangle> rects) :
            this()
        {
            foreach (var r in rects)
            {
                this.Add(r);
            }
        }

        public void Add(Point p)
        {
            if (this.initialized)
            {
                if (p.X < this.min_x)
                {
                    this.min_x = p.X;
                }
                else if (p.X > this.max_x)
                {
                    this.max_x = p.X;
                }
                else
                {
                     // do nothing
                }

                if (p.Y < this.min_y)
                {
                    this.min_y = p.Y;
                    
                }
                else if (p.Y > this.max_y)
                {
                    this.max_y = p.Y;
                }
                else
                {
                    // do nothing
                }
                
            }
            else
            {
                this.min_x = p.X;
                this.max_x = p.X;
                this.min_y = p.Y;
                this.max_y = p.Y;
                this.initialized = true;
            }
        }

        public void Add(Rectangle r)
        {
            this.Add(r.LowerLeft);
            this.Add(r.UpperRight);
        }

        public Rectangle Rectangle
        {
            get
            {
                if (this.HasValue)
                {
                    return new Rectangle(this.min_x,this.min_y,this.max_x,this.max_y);
                }
                else
                {
                    throw new System.ArgumentException("Bounding Box Has no value");
                }
            }
        }

        public bool HasValue => this.initialized;
    }
}