namespace VisioAutomation.Drawing
{
    internal struct ArcSegment
    {
        public readonly double begin;
        public readonly double end;

        internal ArcSegment(double b, double e)
        {
            this.begin = b;
            this.end = e;
        }
    }
}