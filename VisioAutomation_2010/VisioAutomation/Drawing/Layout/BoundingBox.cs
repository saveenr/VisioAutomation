using System.Collections.Generic;

namespace VisioAutomation.Drawing.Layout
{
    public struct BoundingBox
    {
        private bool _initialized;

        private double _min_x;
        private double _min_y;
        private double _max_x;
        private double _max_y;

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
            if (this._initialized)
            {
                if (p.X < this._min_x)
                {
                    this._min_x = p.X;
                }
                else if (p.X > this._max_x)
                {
                    this._max_x = p.X;
                }
                else
                {
                     // do nothing
                }

                if (p.Y < this._min_y)
                {
                    this._min_y = p.Y;
                    
                }
                else if (p.Y > this._max_y)
                {
                    this._max_y = p.Y;
                }
                else
                {
                    // do nothing
                }
                
            }
            else
            {
                this._min_x = p.X;
                this._max_x = p.X;
                this._min_y = p.Y;
                this._max_y = p.Y;
                this._initialized = true;
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
                    return new Rectangle(this._min_x,this._min_y,this._max_x,this._max_y);
                }
                else
                {
                    throw new System.ArgumentException("Bounding Box Has no value");
                }
            }
        }

        public bool HasValue => this._initialized;
    }
}