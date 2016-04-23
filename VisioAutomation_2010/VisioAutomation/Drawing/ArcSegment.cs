namespace VisioAutomation.Drawing
{
    internal struct ArcSegment
    {
        readonly public double begin;
        readonly public double end;

        internal ArcSegment(double b, double e)
        {
            this.begin = b;
            this.end = e;
        }
    }
}