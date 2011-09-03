using IVisio=Microsoft.Office.Interop.Visio;

using IG=InfoGraphicsPy;

namespace InfoGraphicsPy
{
    public struct Point
    {
        private double _x;
        private double _y;

        public Point (double x, double y)
        {
            _x = x;
            _y = y;
        }

        public double X
        {
            get { return _x; }
        }

        public double Y
        {
            get { return _y; }
        }

        public IG.Point Add( double x, double y)
        {
            return new IG.Point(this.X+x,this.Y+y);
        }
    }

    public struct Rectangle
    {
        public double Left { get; private set; }
        public double Bottom { get; private set; }
        public double Right { get; private set; }
        public double Top { get; private set; }

        public Rectangle(double left, double bottom, double right, double top)
            : this()
        {
            if (right < left)
            {
                throw new System.ArgumentException("left must be <=right");
            }

            if (top < bottom)
            {
                throw new System.ArgumentException("bottom must be <= top");
            }

            Left = left;
            Bottom = bottom;
            Right = right;
            Top = top;
        }

        public Rectangle(Point lowerleft, Point upperright)
            : this()
        {
            if (upperright.X < lowerleft.X)
            {
                throw new System.ArgumentException("left must be <=right");
            }

            if (upperright.Y < lowerleft.Y)
            {
                throw new System.ArgumentException("bottom must be <= top");
            }

            Left = lowerleft.X;
            Bottom = lowerleft.Y;
            Right = upperright.X;
            Top = upperright.Y;
        }

        public Rectangle(Point lowerleft, Size s)
            : this()
        {
            if (s.Width < 0)
            {
                throw new System.ArgumentOutOfRangeException("s", "width must be non-negative");
            }

            if (s.Height < 0)
            {
                throw new System.ArgumentOutOfRangeException("s", "height must be non-negative");
            }

            Left = lowerleft.X;
            Bottom = lowerleft.Y;
            Right = lowerleft.X + s.Width;
            Top = lowerleft.Y + s.Height;
        }

        public static Rectangle FromCenterPoint(double x, double y, double w, double h)
        {
            if (w < 0)
            {
                throw new System.ArgumentOutOfRangeException("w", "width must be non-negative");
            }

            if (h < 0)
            {
                throw new System.ArgumentOutOfRangeException("h", "height must be non-negative");
            }

            var xradius = w / 2.0;
            var yradius = h / 2.0;
            var r = new Rectangle(x - xradius, y - yradius, x + xradius, y + yradius);
            return r;
        }

        public static Rectangle FromCenterPoint(Point p, double width, double height)
        {
            return FromCenterPoint(p.X, p.Y, width, height);
        }

        public override string ToString()
        {
            string s = string.Format(System.Globalization.CultureInfo.InvariantCulture, "({0:0.#####},{1:0.#####},{2:0.#####},{3:0.#####})",
                                     Left, Bottom, Right, Top);
            return s;
        }

        public Point LowerLeft
        {
            get { return new Point(Left, Bottom); }
        }

        public Point LowerRight
        {
            get { return new Point(Right, Bottom); }
        }

        public Point UpperLeft
        {
            get { return new Point(Left, Top); }
        }

        public Point UpperRight
        {
            get { return new Point(Right, Top); }
        }

        public Size Size
        {
            get { return new Size(Width, Height); }
        }

        public double Width
        {
            get { return Right - Left; }
        }

        public double Height
        {
            get { return Top - Bottom; }
        }

        public Point Center
        {
            get { return new Point((Left + Right) / 2.0, (Bottom + Top) / 2.0); }
        }

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
            var r2 = new Rectangle(Left + dx, Bottom + dy, Right + dx, Top + dy);
            return r2;
        }

        public Rectangle Add(Size s)
        {
            var r2 = new Rectangle(Left + s.Width, Bottom + s.Height, Right + s.Width, Top + s.Height);
            return r2;
        }

        public Rectangle Add(Point s)
        {
            var r2 = new Rectangle(Left + s.X, Bottom + s.Y, Right + s.X, Top + s.Y);
            return r2;
        }


        public Rectangle Subtract(double dx, double dy)
        {
            var r2 = new Rectangle(Left - dx, Bottom - dy, Right - dx, Top - dy);
            return r2;
        }

        public Rectangle Subtract(Size s)
        {
            var r2 = new Rectangle(Left - s.Width, Bottom - s.Height, Right - s.Width, Top - s.Height);
            return r2;
        }

        public Rectangle Subtract(Point s)
        {
            var r2 = new Rectangle(Left - s.X, Bottom - s.Y, Right - s.X, Top - s.Y);
            return r2;
        }


        public Rectangle Multiply(double sx, double sy)
        {
            var r2 = new Rectangle(Left * sx, Bottom * sy, Right * sx, Top * sy);
            return r2;
        }
    }

    public struct Size
    {
        public double Width { get; private set; }
        public double Height { get; private set; }

        public Size(double width, double height)
            : this()
        {
            if (width < 0.0)
            {
                throw new System.ArgumentOutOfRangeException("width");
            }
            if (height < 0.0)
            {
                throw new System.ArgumentOutOfRangeException("height");
            }
            Width = width;
            Height = height;
        }


        public override string ToString()
        {
            string s = string.Format(System.Globalization.CultureInfo.InvariantCulture, "({0:0.#####}, {1:0.#####})", Width, Height);
            return s;
        }

        public Size Multiply(double amount)
        {
            return new Size(Width * amount, Height * amount);
        }

        public Size Multiply(double width, double height)
        {
            return new Size(Width * width, Height * height);
        }

        public Size Divide(double amount)
        {
            return new Size(Width / amount, Height / amount);
        }

        public static Size operator *(Size left_size, double right_size)
        {
            return left_size.Multiply(right_size);
        }

        public static Size operator /(Size left_size, double right_size)
        {
            return left_size.Divide(right_size);
        }

        public Size Add(Point point)
        {
            return new Size(Width + point.X, Height + point.Y);
        }

        public Size Add(Size size)
        {
            return new Size(Width + size.Width, Height + size.Height);
        }

        public Size Add(double width, double height)
        {
            return new Size(Width + width, Height + height);
        }

        public static Size operator +(Size size, Point point)
        {
            return size.Add(point);
        }

        public static Size operator +(Size left_size, Size right_size)
        {
            return left_size.Add(right_size);
        }
    }

    public static class DrawUtil
    {
        public static IVisio.Shape DrawCircleFromCenter(IVisio.Page page, IG.Point center, double r)
        {
            var lowerleft = center.Add(-r, -r);
            var upperright = center.Add(r, r);
            var shape = page.DrawOval( lowerleft.X, lowerleft.Y, upperright.X, upperright.Y);
            return shape;
        }

        public static IVisio.Shape DrawCircleFromCenter(IVisio.Page page, double x, double y, double r)
        {
            var shape = page.DrawOval(x - r, y - r, x + r, y + r);
            return shape;
        }
    }
}
