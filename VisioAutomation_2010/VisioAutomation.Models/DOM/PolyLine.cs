using System.Collections.Generic;
using System.Linq;

namespace VisioAutomation.Models.Dom
{
    public class PolyLine : BaseShape
    {
        public List<VisioAutomation.Core.Point> Points { get; }

        public PolyLine(params double[] doubles)
        {
            this.Points = VisioAutomation.Core.Point.FromDoubles(doubles).ToList();
        }

        public PolyLine(IEnumerable<VisioAutomation.Core.Point> pts)
        {
            this.Points = pts.ToList();
        }
    }
}