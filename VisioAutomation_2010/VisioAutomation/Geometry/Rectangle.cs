namespace VisioAutomation.Geometry
{
    public struct Rectangle
    {
        public double Left { get; }
        public double Bottom { get; }
        public double Right { get; }
        public double Top { get; }

        public Rectangle(double left, double bottom, double right, double top)
            : this()
        {
            if (right < left)
            {
                throw new System.ArgumentException("left must be <= right");
            }

            if (top < bottom)
            {
                throw new System.ArgumentException("bottom must be <= top");
            }

            this.Left = left;
            this.Bottom = bottom;
            this.Right = right;
            this.Top = top;
        }

        public Rectangle(Point bottomleft, Point topright)
            : this(bottomleft.X, bottomleft.Y, topright.X, topright.Y)
        {
        }

        public Rectangle(Point bottomleft, Size s)
            : this()
        {
            if (s.Width < 0)
            {
                throw new System.ArgumentOutOfRangeException(nameof(s), "width must be non-negative");
            }

            if (s.Height < 0)
            {
                throw new System.ArgumentOutOfRangeException(nameof(s), "height must be non-negative");
            }

            this.Left = bottomleft.X;
            this.Bottom = bottomleft.Y;
            this.Right = bottomleft.X + s.Width;
            this.Top = bottomleft.Y + s.Height;
        }

        public static Rectangle FromCenterPoint(double x, double y, double w, double h)
        {
            if (w < 0)
            {
                throw new System.ArgumentOutOfRangeException(nameof(w), "width must be non-negative");
            }

            if (h < 0)
            {
                throw new System.ArgumentOutOfRangeException(nameof(h), "height must be non-negative");
            }

            var xradius = w/2.0;
            var yradius = h/2.0;
            var r = new Rectangle(x - xradius, y - yradius, x + xradius, y + yradius);
            return r;
        }

        public static Rectangle FromCenterPoint(Point p, double width, double height)
        {
            return Rectangle.FromCenterPoint(p.X, p.Y, width, height);
        }

        public override string ToString()
        {
            string s = string.Format(System.Globalization.CultureInfo.InvariantCulture, "({0:0.#####},{1:0.#####},{2:0.#####},{3:0.#####})", this.Left, this.Bottom, this.Right, this.Top);
            return s;
        }

        public Point LowerLeft => new Point(this.Left, this.Bottom);

        public Point LowerRight => new Point(this.Right, this.Bottom);

        public Point UpperLeft => new Point(this.Left, this.Top);

        public Point UpperRight => new Point(this.Right, this.Top);

        public Size Size => new Size(this.Width, this.Height);

        public double Width => this.Right - this.Left;

        public double Height => this.Top - this.Bottom;

        public Point Center => new Point((this.Left + this.Right)/2.0, (this.Bottom + this.Top)/2.0);

        public static Rectangle operator +(Rectangle r, Point p)
        {
            return r.Add(p.X, p.Y);
        }

        public static Rectangle operator -(Rectangle r, Point p)
        {
            return r.Subtract(p.X, p.Y);
        }

        public static Rectangle operator *(Rectangle r, double s)
        {
            return r.Multiply(s, s);
        }

        public Rectangle Add(double dx, double dy)
        {
            var r2 = new Rectangle(this.Left + dx, this.Bottom + dy, this.Right + dx, this.Top + dy);
            return r2;
        }

        public Rectangle Add(Size s)
        {
            var r2 = new Rectangle(this.Left + s.Width, this.Bottom + s.Height, this.Right + s.Width, this.Top + s.Height);
            return r2;
        }

        public Rectangle Add(Point s)
        {
            var r2 = new Rectangle(this.Left + s.X, this.Bottom + s.Y, this.Right + s.X, this.Top + s.Y);
            return r2;
        }


        public Rectangle Subtract(double dx, double dy)
        {
            var r2 = new Rectangle(this.Left - dx, this.Bottom - dy, this.Right - dx, this.Top - dy);
            return r2;
        }

        public Rectangle Subtract(Size s)
        {
            var r2 = new Rectangle(this.Left - s.Width, this.Bottom - s.Height, this.Right - s.Width, this.Top - s.Height);
            return r2;
        }

        public Rectangle Subtract(Point s)
        {
            var r2 = new Rectangle(this.Left - s.X, this.Bottom - s.Y, this.Right - s.X, this.Top - s.Y);
            return r2;
        }


        public Rectangle Multiply(double sx, double sy)
        {
            var r2 = new Rectangle(this.Left*sx, this.Bottom*sy, this.Right*sx, this.Top*sy);
            return r2;
        }
    }
}