using System.Collections.Generic;
using VA=VisioAutomation;

namespace VisioAutomation.Drawing
{
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

        public Rectangle(VA.Drawing.Point lowerleft, Drawing.Point upperright)
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

        public Rectangle(VA.Drawing.Point lowerleft, VA.Drawing.Size s)
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

            var xradius = w/2.0;
            var yradius = h/2.0;
            var r = new Rectangle(x - xradius, y - yradius, x + xradius, y + yradius);
            return r;
        }

        public static Rectangle FromCenterPoint(VA.Drawing.Point p, double width, double height)
        {
            return FromCenterPoint(p.X, p.Y, width, height);
        }

        public override string ToString()
        {
            string s = string.Format(System.Globalization.CultureInfo.InvariantCulture, "({0:0.#####},{1:0.#####},{2:0.#####},{3:0.#####})",
                                     Left, Bottom, Right, Top);
            return s;
        }

        public VA.Drawing.Point LowerLeft
        {
            get { return new Drawing.Point(Left, Bottom); }
        }

        public VA.Drawing.Point LowerRight
        {
            get { return new Drawing.Point(Right, Bottom); }
        }

        public VA.Drawing.Point UpperLeft
        {
            get { return new VA.Drawing.Point(Left, Top); }
        }

        public VA.Drawing.Point UpperRight
        {
            get { return new VA.Drawing.Point(Right, Top); }
        }

        public VA.Drawing.Size Size
        {
            get { return new VA.Drawing.Size(Width, Height); }
        }

        public double Width
        {
            get { return Right - Left; }
        }

        public double Height
        {
            get { return Top - Bottom; }
        }

        public VA.Drawing.Point Center
        {
            get { return new Drawing.Point((Left + Right)/2.0, (Bottom + Top)/2.0); }
        }

        public static Rectangle operator +(Rectangle r, VA.Drawing.Point p)
        {
            return r.Add(p.X, p.Y);
        }

        public static Rectangle operator -(Rectangle r, VA.Drawing.Point p)
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
            var r2 = new Rectangle(Left*sx, Bottom*sy, Right*sx, Top*sy);
            return r2;
        }
    }
}