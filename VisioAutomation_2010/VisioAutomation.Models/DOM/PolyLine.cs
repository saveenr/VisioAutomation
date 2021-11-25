using System.Collections.Generic;
using System.Linq;

namespace VisioAutomation.Models.Dom;

public class PolyLine : BaseShape
{
    public List<VisioAutomation.Geometry.Point> Points { get; }

    public PolyLine(params double[] doubles)
    {
        this.Points = VisioAutomation.Geometry.Point.FromDoubles(doubles).ToList();
    }

    public PolyLine(IEnumerable<VisioAutomation.Geometry.Point> pts)
    {
        this.Points = pts.ToList();
    }
}