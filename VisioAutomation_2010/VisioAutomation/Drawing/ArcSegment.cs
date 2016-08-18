namespace VisioAutomation.Drawing
{
    internal struct ArcSegment
    {
        public readonly double Begin;
        public readonly double End;

        internal ArcSegment(double b, double e)
        {
            this.Begin = b;
            this.End = e;
        }
    }
}