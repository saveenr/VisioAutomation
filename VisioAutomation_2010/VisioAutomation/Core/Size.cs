namespace VisioAutomation.Core
{
    public struct Size
    {
        public double Width { get; }
        public double Height { get; }

        public Size(double width, double height)
            : this()
        {
            if (width < 0.0)
            {
                throw new System.ArgumentOutOfRangeException(nameof(width));
            }
            if (height < 0.0)
            {
                throw new System.ArgumentOutOfRangeException(nameof(height));
            }
            this.Width = width;
            this.Height = height;
        }
        
        public override string ToString()
        {
            var culture = System.Globalization.CultureInfo.InvariantCulture;
            string s = string.Format(culture, "({0:0.#####}, {1:0.#####})", this.Width, this.Height);
            return s;
        }

        public static Size operator +(Size left, Point right) => left.Add(right);
        public static Size operator +(Size left, Size right) => left.Add(right);
        public static Size operator -(Size left, Point right) => left.Subtract(right);
        public static Size operator -(Size left, Size right) => left.Subtract(right);
        public static Size operator *(Size left, Point right) => left.Multiply(right);
        public static Size operator *(Size left, Size right) => left.Multiply(right);
        public static Size operator /(Size left, Point right) => left.Divide(right);
        public static Size operator /(Size left, Size right) => left.Divide(right);

        public Size Multiply(double sx, double sy) => new Size(this.Width*sx, this.Height*sy);
        public Size Multiply(Size s) => this.Multiply(s.Width, s.Height);
        public Size Multiply(Point p) => this.Multiply(p.X, p.Y);

        public Size Divide(double sx, double sy) => new Size(this.Width / sx, this.Height / sy);
        public Size Divide(Size s) => this.Divide(s.Width, s.Height);
        public Size Divide(Point p) => this.Divide(p.X, p.Y);

        public Size Add(double dx, double dy) => new Size(this.Width + dx, this.Height + dy);
        public Size Add(Size s) => this.Add(s.Width, s.Height);
        public Size Add(Point p) => this.Add(p.X, p.Y);

        public Size Subtract(double dx, double dy) => new Size(this.Width - dx, this.Height - dy);
        public Size Subtract(Size s) => this.Subtract(s.Width, s.Height);
        public Size Subtract(Point p) => this.Subtract(p.X, p.Y);
    }
}