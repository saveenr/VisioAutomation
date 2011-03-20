using VA=VisioAutomation;

namespace VisioAutomation.Drawing
{
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
            return new Size(Width*amount, Height*amount);
        }

        public Size Multiply(double width, double height)
        {
            return new Size(Width*width, Height*height);
        }

        public Size Divide(double amount)
        {
            return new Size(Width/amount, Height/amount);
        }

        public static Size operator *(Size left_size, double right_size)
        {
            return left_size.Multiply(right_size);
        }

        public static Size operator /(Size left_size, double right_size)
        {
            return left_size.Divide(right_size);
        }

        public Size Add(VA.Drawing.Point point)
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

        public static Size operator +(Size size, VA.Drawing.Point point)
        {
            return size.Add(point);
        }

        public static Size operator +(Size left_size, Size right_size)
        {
            return left_size.Add(right_size);
        }
    }
}