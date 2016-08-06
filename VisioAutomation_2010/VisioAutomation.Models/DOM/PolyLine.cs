using System.Collections.Generic;
using System.Linq;

namespace VisioAutomation.DOM
{
    public class PolyLine : BaseShape
    {
        public List<Drawing.Point> Points { get; private set; }

        public PolyLine(params double[] doubles)
        {
            this.Points = Drawing.Point.FromDoubles(doubles).ToList();
        }

        public PolyLine(IEnumerable<Drawing.Point> pts)
        {
            this.Points = pts.ToList();
        }
    }
}