namespace VisioAutomationMin
{
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

        public double Width
        {
            get { return this.Right - this.Left; }
        }

        public double Height
        {
            get { return this.Top - this.Bottom; }
        }

        public override string ToString()
        {
            return string.Format("({0},{1},{2},{3}", this.Left, this.Bottom, this.Right, this.Top);
        }

        public Point Center
        {
            get { 
                double x = this.Left + this.Width / 2.0;
                double y = this.Bottom + this.Height / 2.0;
                var p = new Point(x, y);
                return p;
            }
        }
    }
}