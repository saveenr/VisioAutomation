using System.Collections.Generic;
using System.Linq;

namespace VisioAutomation.Models.DOM
{
    public class BezierCurve : BaseShape
    {
        public List<Drawing.Point> ControlPoints { get; private set; }
        public int Degree { get; private set; }

        public BezierCurve(IEnumerable<Drawing.Point> pts)
        {
            this.Degree = 3;
            this.ControlPoints = pts.ToList();
        }

        public BezierCurve(IEnumerable<double> pts)
        {
            this.Degree = 3;
            this.ControlPoints = Drawing.Point.FromDoubles(pts).ToList();
        }
    }
}