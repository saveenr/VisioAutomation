using System.Collections.Generic;

namespace VisioAutomation.Geometry
{
    public struct Point
    {
        public double X { get; }
        public double Y { get; }

        public Point(double x, double y)
            : this()
        {
            this.X = x;
            this.Y = y;
        }

        public override string ToString()
        {
            var invariant_culture = System.Globalization.CultureInfo.InvariantCulture;
            string s = string.Format(invariant_culture, "Point({0:0.#####}, {1:0.#####})", this.X, this.Y);
            return s;
        }

        public static Point operator -(Point pa, Point pb) => pa.Subtract(pb);
        public static Point operator +(Point pa, Point pb) => pa.Add(pb);
        public static Point operator *(Point pa, double s) => pa.Multiply(s, s);
        public static Point operator *(Point pa, Size s) => pa.Multiply(s);
        public static Point operator *(Point pa, Point pb) => pa.Multiply(pb.X, pb.Y);

        public Point Add(double dx, double dy) => new Point(this.X + dx, this.Y + dy);
        public Point Add(Point p) => new Point(this.X + p.X, this.Y + p.Y);
        public Point Add(Size s) => new Point(this.X + s.Width, this.Y + s.Height);

        public Point Subtract(double dx, double dy) => new Point(this.X - dx, this.Y - dy);
        public Point Subtract(Size s) => new Point(this.X - s.Width, this.Y - s.Height);
        public Point Subtract(Point p) => new Point(this.X - p.X, this.Y - p.Y);

        public Point Multiply(double sx, double sy) => new Point(this.X*sx, this.Y*sy);
        public Point Multiply(Size s) => new Point(this.X*s.Width, this.Y*s.Height);

        public Point Divide(double sx, double sy) => new Point(this.X/sx, this.Y/sy);

        public static IEnumerable<Point> FromDoubles(IEnumerable<double> doubles)
        {
            if (doubles == null)
            {
                throw new System.ArgumentNullException(nameof(doubles));
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
                    yield return new Point(even_value, value);
                }
                count++;
            }
        }

        public static IEnumerable<double> ToDoubles(IEnumerable<Point> points)
        {
            foreach (var p in points)
            {
                yield return p.X;
                yield return p.Y;
            }
        }
    }
}