namespace VisioAutomation.Models.Geometry
{
    public struct LineSegment
    {
        public VisioAutomation.Core.Point Start { get; }
        public VisioAutomation.Core.Point End { get; }

        public LineSegment(VisioAutomation.Core.Point start, VisioAutomation.Core.Point end)
        {
            this.Start = start;
            this.End = end;
        }
    }
}