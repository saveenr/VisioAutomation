namespace VisioAutomation.Geometry
{
    public struct LineSegment
    {
        public Point Start { get; }
        public Point End { get; }

        public LineSegment(Point start, Point end)
        {
            this.Start = start;
            this.End = end;
        }
    }
}