using IVisio = Microsoft.Office.Interop.Visio;

namespace VisUtil
{
    public struct Point
    {
        public readonly double X;
        public readonly double Y;

        public Point(double x, double y)
        {
            this.X = x;
            this.Y = y;
        }

        public static Point operator +(Point p0, Point p1)
        {
            return new Point(p0.X + p1.X, p0.Y + p1.Y);
        }

        public static Point operator +(Point p0, Size p1)
        {
            return new Point(p0.X + p1.Width, p0.Y + p1.Height);
        }


        public static Point operator -(Point p0, Point p1)
        {
            return new Point(p0.X - p1.X, p0.Y-p1.Y);
        }
    }

    public struct Size
    {
        public readonly double Width;
        public readonly double Height;

        public Size(double width, double height)
        {
            this.Width = width;
            this.Height = height;
        }

        public static Size operator +(Size p0, Size p1)
        {
            return new Size(p0.Width + p1.Width, p0.Height + p1.Height);
        }

        public static Size operator -(Size p0, Size p1)
        {
            return new Size(p0.Width - p1.Width, p0.Height - p1.Height);
        }
    }

    public struct Rectangle
    {
        public readonly double Left;
        public readonly double Bottom;
        public readonly double Right;
        public readonly double Top;

        public Rectangle(double left, double bottom, double right, double top)
        {
            this.Left = left;
            this.Bottom = bottom;
            this.Right = right;
            this.Top = top;
        }

        public Rectangle(Point lowerleft, Point upperright)
        {
            this.Left = lowerleft.X;
            this.Bottom = lowerleft.Y;
            this.Right = upperright.X;
            this.Top = upperright.Y;
        }

        public Rectangle(Point lowerleft, Size size)
        {
            this.Left = lowerleft.X;
            this.Bottom = lowerleft.Y;
            this.Right = lowerleft.X + size.Width;
            this.Top = lowerleft.Y + size.Height;
        }

        public double Width
        {
            get { return this.Right - this.Left; }
        }

        public double Height
        {
            get { return this.Top - this.Bottom; }
        }
        
        public Point LowerLeft
        {
            get
            {
                return new Point(this.Left,this.Bottom);
            }
        }

        public Point UpperLeft
        {
            get
            {
                return new Point(this.Left, this.Top);
            }
        }

        public Point LowerRight
        {
            get
            {
                return new Point(this.Right, this.Bottom);
            }
        }

        public Point UpperRight
        {
            get
            {
                return new Point(this.Right, this.Top);
            }
        }
    }
}
