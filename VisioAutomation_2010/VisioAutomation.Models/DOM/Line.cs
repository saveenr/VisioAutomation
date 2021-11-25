namespace VisioAutomation.Models.Dom;

public class Line : BaseShape
{
    public VisioAutomation.Geometry.Point P0 { get; }
    public VisioAutomation.Geometry.Point P1 { get; }

    public Line(double x0, double y0, double x1, double y1)
    {
        this.P0 = new VisioAutomation.Geometry.Point(x0, y0);
        this.P1 = new VisioAutomation.Geometry.Point(x1, y1);
    }

    public Line(VisioAutomation.Geometry.Point p0, VisioAutomation.Geometry.Point p1)
    {
        this.P0 = p0;
        this.P1 = p1;
    }
}