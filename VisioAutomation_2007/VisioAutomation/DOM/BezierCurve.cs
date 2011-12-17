using System.Collections.Generic;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using System.Linq;

namespace VisioAutomation.DOM
{
    public class BezierCurve : Shape
    {
        public List<VA.Drawing.Point> ControlPoints { get; private set; }
        public int Degree { get; private set; }

        public BezierCurve(IEnumerable<VA.Drawing.Point> pts)
        {
            Degree = 3;
            this.ControlPoints = pts.ToList();
        }

        public BezierCurve(IEnumerable<double> pts)
        {
            Degree = 3;
            this.ControlPoints = VA.Drawing.Point.FromDoubles(pts).ToList();
        }
    }
}