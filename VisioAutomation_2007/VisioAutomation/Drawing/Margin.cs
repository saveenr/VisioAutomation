using VA=VisioAutomation;

namespace VisioAutomation.Drawing
{
    public struct Margin
    {
        public double Left { get; private set; }
        public double Bottom { get; private set; }
        public double Right { get; private set; }
        public double Top { get; private set; }

        public Margin(double left, double bottom, double right, double top)
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

        public override string ToString()
        {
            string s = string.Format(System.Globalization.CultureInfo.InvariantCulture, "({0:0.#####},{1:0.#####},{2:0.#####},{3:0.#####})",
                                     Left, Bottom, Right, Top);
            return s;
        }

        public double TotalWidth
        {
            get { return Right + Left; }
        }

        public double TotalHeight
        {
            get { return Top + Bottom; }
        }

    }
}