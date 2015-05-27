namespace VisioAutomation.Drawing
{
    public struct Margin
    {
        public double Left { get; }
        public double Bottom { get; }
        public double Right { get; }
        public double Top { get; }

        public Margin(double left, double bottom, double right, double top)
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

        public override string ToString()
        {
            string s = string.Format(System.Globalization.CultureInfo.InvariantCulture, "({0:0.#####},{1:0.#####},{2:0.#####},{3:0.#####})", this.Left, this.Bottom, this.Right, this.Top);
            return s;
        }

        public double TotalWidth => this.Right + this.Left;

        public double TotalHeight => this.Top + this.Bottom;
    }
}