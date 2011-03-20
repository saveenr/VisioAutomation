using System.Collections.Generic;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using System.Linq;

namespace VisioAutomation.DOM
{
    public class PolyLine : Shape
    {
        public List<VA.Drawing.Point> Points { get; private set; }

        public PolyLine(params double[] doubles)
        {
            this.Points = VA.Drawing.DrawingUtil.DoublesToPoints(doubles).ToList();
        }

        public PolyLine(IEnumerable<VA.Drawing.Point> pts)
        {
            this.Points = pts.ToList();
        }
    }
}