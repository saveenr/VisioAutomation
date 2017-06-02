using System.Collections.Generic;
using System.Linq;

namespace VisioAutomation.Models.Dom
{
    public class PolyLine : BaseShape
    {
        public List<Geometry.Point> Points { get; private set; }

        public PolyLine(params double[] doubles)
        {
            this.Points = Geometry.Point.FromDoubles(doubles).ToList();
        }

        public PolyLine(IEnumerable<Geometry.Point> pts)
        {
            this.Points = pts.ToList();
        }
    }
}