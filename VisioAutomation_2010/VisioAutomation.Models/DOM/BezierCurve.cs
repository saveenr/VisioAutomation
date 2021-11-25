using System.Collections.Generic;
using System.Linq;

namespace VisioAutomation.Models.Dom;

public class BezierCurve : BaseShape
{
    public List<VisioAutomation.Geometry.Point> ControlPoints { get; }
    public int Degree { get; }

    public BezierCurve(IEnumerable<VisioAutomation.Geometry.Point> pts)
    {
        this.Degree = 3;
        this.ControlPoints = pts.ToList();
    }

    public BezierCurve(IEnumerable<double> pts)
    {
        this.Degree = 3;
        this.ControlPoints = VisioAutomation.Geometry.Point.FromDoubles(pts).ToList();
    }
}