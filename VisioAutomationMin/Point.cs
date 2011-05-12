namespace VisioAutomationMin
{
    public struct Point
    {
        public readonly double X;
        public readonly double Y;

        public Point(double x, double y)
        {
            this.X = x;
            this.Y = y;
        }

        public Point Add(double x, double y)
        {
            return new Point(this.X+x, this.Y+y);
        }

        public Point Add(Point p)
        {
            return new Point(this.X + p.X, this.Y + p.Y);
        }

    }
}