using System.Collections.Generic;
using System.Linq;

namespace VisioAutomation.Models.Dom
{
    public class BezierCurve : BaseShape
    {
        public List<Geometry.Point> ControlPoints { get; private set; }
        public int Degree { get; private set; }

        public BezierCurve(IEnumerable<Geometry.Point> pts)
        {
            this.Degree = 3;
            this.ControlPoints = pts.ToList();
        }

        public BezierCurve(IEnumerable<double> pts)
        {
            this.Degree = 3;
            this.ControlPoints = Geometry.Point.FromDoubles(pts).ToList();
        }
    }
}