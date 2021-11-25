namespace VisioAutomation.Models.Geometry;

internal struct ArcSegment
{
    public readonly double Begin;
    public readonly double End;

    internal ArcSegment(double begin, double end)
    {
        this.Begin = begin;
        this.End = end;
    }
}