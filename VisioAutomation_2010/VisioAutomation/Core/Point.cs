using System.Collections.Generic;

namespace VisioAutomation.Core
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
            var culture = System.Globalization.CultureInfo.InvariantCulture;
            string s = string.Format(culture, "Point({0:0.#####}, {1:0.#####})", this.X, this.Y);
            return s;
        }

        public static Point operator -(Point left, Point right) => left.Subtract(right);
        public static Point operator -(Point left, Size right) => left.Subtract(right);
        public static Point operator +(Point left, Point right) => left.Add(right);
        public static Point operator +(Point left, Size right) => left.Add(right);
        public static Point operator *(Point left, Point right) => left.Multiply(right);
        public static Point operator *(Point left, Size right) => left.Multiply(right);
        public static Point operator /(Point left, Point right) => left.Divide(right);
        public static Point operator /(Point left, Size right) => left.Divide(right);

        public Point Add(double dx, double dy) => new Point(this.X + dx, this.Y + dy);
        public Point Add(Size s) => this.Add(s.Width, s.Height);
        public Point Add(Point p) => this.Add(p.X, p.Y);

        public Point Subtract(double dx, double dy) => new Point(this.X - dx, this.Y - dy);
        public Point Subtract(Size s) => this.Subtract(s.Width, s.Height);
        public Point Subtract(Point p) => this.Subtract(p.X, p.Y);

        public Point Multiply(double sx, double sy) => new Point(this.X*sx, this.Y*sy);
        public Point Multiply(Size s) => this.Multiply(s.Width, s.Height);
        public Point Multiply(Point p) => this.Multiply(p.X, p.Y);

        public Point Divide(double sx, double sy) => new Point(this.X/sx, this.Y/sy);
        public Point Divide(Size s) => this.Divide(s.Width, s.Height);
        public Point Divide(Point p) => this.Divide(p.X, p.Y);

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