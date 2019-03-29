namespace VisioAutomation.Models.Geometry
{
    public struct LineSegment
    {
        public VisioAutomation.Geometry.Point Start { get; }
        public VisioAutomation.Geometry.Point End { get; }

        public LineSegment(VisioAutomation.Geometry.Point start, VisioAutomation.Geometry.Point end)
        {
            this.Start = start;
            this.End = end;
        }
    }
}