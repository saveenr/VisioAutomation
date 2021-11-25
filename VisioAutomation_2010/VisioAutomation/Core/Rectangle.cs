namespace VisioAutomation.Core
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

        public override string ToString()
        {
            var culture = System.Globalization.CultureInfo.InvariantCulture;
            string s = string.Format(culture, "({0:0.#####},{1:0.#####},{2:0.#####},{3:0.#####})", this.Left, this.Bottom, this.Right, this.Top);
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

        public static Rectangle operator +(Rectangle left, Point right) => left.Add(right);
        public static Rectangle operator +(Rectangle left, Size right) => left.Add(right);
        public static Rectangle operator -(Rectangle left, Point right) => left.Subtract(right);
        public static Rectangle operator -(Rectangle left, Size right) => left.Subtract(right);
        public static Rectangle operator *(Rectangle left, Point right) => left.Multiply(right);
        public static Rectangle operator *(Rectangle left, Size right) => left.Multiply(right);
        public static Rectangle operator /(Rectangle left, Point right) => left.Divide(right);
        public static Rectangle operator /(Rectangle left, Size right) => left.Divide(right);

        public Rectangle Add(double dx, double dy) => new Rectangle(this.Left + dx, this.Bottom + dy, this.Right + dx, this.Top + dy);
        public Rectangle Add(Size s) => this.Add(s.Width,s.Height);
        public Rectangle Add(Point p) => this.Add(p.X, p.Y);

        public Rectangle Subtract(double dx, double dy) => new Rectangle(this.Left - dx, this.Bottom - dy, this.Right - dx, this.Top - dy);
        public Rectangle Subtract(Size s) => this.Subtract(s.Width, s.Height);
        public Rectangle Subtract(Point p) => this.Subtract(p.X, p.Y);

        public Rectangle Multiply(double sx, double sy) => new Rectangle(this.Left*sx, this.Bottom*sy, this.Right*sx, this.Top*sy);
        public Rectangle Multiply(Size s) => this.Multiply(s.Width, s.Height);
        public Rectangle Multiply(Point p) => this.Multiply(p.X, p.Y);

        public Rectangle Divide(double sx, double sy) => new Rectangle(this.Left / sx, this.Bottom / sy, this.Right / sx, this.Top / sy);
        public Rectangle Divide(Size s) => this.Divide(s.Width, s.Height);
        public Rectangle Divide(Point p) => this.Divide(p.X, p.Y);

    }
}