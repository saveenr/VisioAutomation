namespace VisioAutomation.Geometry
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
            string s = string.Format(System.Globalization.CultureInfo.InvariantCulture, "({0:0.#####}, {1:0.#####})", this.Width, this.Height);
            return s;
        }

        public Size Multiply(double amount) => new Size(this.Width*amount, this.Height*amount);
        public Size Multiply(double width, double height) => new Size(this.Width*width, this.Height*height);

        public Size Divide(double amount) => new Size(this.Width/amount, this.Height/amount);

        public Size Add(Point point) => new Size(this.Width + point.X, this.Height + point.Y);
        public Size Add(Size size) => new Size(this.Width + size.Width, this.Height + size.Height);
        public Size Add(double width, double height) => new Size(this.Width + width, this.Height + height);

        public static Size operator +(Size size, Point point) => size.Add(point);
        public static Size operator +(Size left_size, Size right_size) => left_size.Add(right_size);
        public static Size operator *(Size left_size, double right_size) => left_size.Multiply(right_size);
        public static Size operator /(Size left_size, double right_size) => left_size.Divide(right_size);
    }
}