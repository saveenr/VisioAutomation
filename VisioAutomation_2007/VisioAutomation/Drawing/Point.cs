using System.Collections.Generic;
using VA=VisioAutomation;

namespace VisioAutomation.Drawing
{
    public struct Point
    {
        public double X { get; private set; }
        public double Y { get; private set; }

        public Point(double x, double y)
            : this()
        {
            X = x;
            Y = y;
        }


        public override string ToString()
        {
            var invariant_culture = System.Globalization.CultureInfo.InvariantCulture;
            string s = string.Format(invariant_culture, "Point({0:0.#####}, {1:0.#####})", X, Y);
            return s;
        }

        public static Point operator -(Point pa, Point pb)
        {
            var result = new Point(pa.X - pb.X, pa.Y - pb.Y);
            return result;
        }

        public static Point operator +(Point pa, Point pb)
        {
            var result = new Point(pa.X + pb.X, pa.Y + pb.Y);
            return result;
        }

        public static Point operator *(Point pa, double s)
        {
            var result = new Point(pa.X*s, pa.Y*s);
            return result;
        }

        public static Point operator *(Point pa, VA.Drawing.Size s)
        {
            var result = new Point(pa.X*s.Width, pa.Y*s.Height);
            return result;
        }

        public Point Add(double dx, double dy)
        {
            var new_point = new Point(X + dx, Y + dy);
            return new_point;
        }

        public Point Subtract(double dx, double dy)
        {
            var new_point = new Point(X - dx, Y - dy);
            return new_point;
        }

        public Point Add(Point p)
        {
            var new_point = new Point(X + p.X, Y + p.Y);
            return new_point;
        }

        public Point Subtract(Point p)
        {
            var new_point = new Point(X - p.X, Y - p.Y);
            return new_point;
        }

        public Point Add(VA.Drawing.Size s)
        {
            var new_point = new Point(X + s.Width, Y + s.Height);
            return new_point;
        }

        public Point Subtract(VA.Drawing.Size s)
        {
            var new_point = new Point(X - s.Width, Y - s.Height);
            return new_point;
        }

        public static Point operator *(Point pa, Point pb)
        {
            return pa.Multiply(pb.X, pb.Y);
        }

        public Point Multiply(double s)
        {
            return Multiply(s, s);
        }

        public Point Multiply(double sx, double sy)
        {
            var new_point = new Point(X*sx, Y*sy);
            return new_point;
        }

        public Point Multiply(VA.Drawing.Size s)
        {
            var new_point = new Point(X*s.Width, Y*s.Height);
            return new_point;
        }

        public Point Divide(double sx, double sy)
        {
            var new_point = new Point(X/sx, Y/sy);
            return new_point;
        }

        public Point Divide(double s)
        {
            return Divide(s, s);
        }

        public static IEnumerable<Drawing.Point> FromDoubles(IEnumerable<double> doubles)
        {
            if (doubles == null)
            {
                throw new System.ArgumentNullException("doubles");
            }

            int count = 0;
            double even_value = default(double);
            foreach (var value in doubles)
            {
                if ((count % 2) == 0)
                {
                    even_value = value;
                }
                else
                {
                    yield return new Drawing.Point(even_value, value);
                }
                count++;
            }
        }

        public static IEnumerable<double> ToDoubles(IEnumerable<Drawing.Point> points)
        {
            foreach (var p in points)
            {
                yield return p.X;
                yield return p.Y;
            }
        }

    }
}