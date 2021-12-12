using System.Linq;
using VisioAutomation.Extensions;

namespace VSamples.Samples.Misc
{
    public class BezierSimple : SampleMethodBase
    {
        public override void RunSample()
        {
            // Draw a Simple Bezier curve

            var page = SampleEnvironment.Application.ActiveDocument.Pages.Add();

            var points = new[] {0.0, 0.0, 1.0, 2.0, 3.0, 0.5, 4.0, 0.5};
            var cpoints = VisioAutomation.Core.Point.FromDoubles(points).ToList();
            var s0 = page.DrawBezier(cpoints);
            s0.Text = "Bezier curve";
            foreach (var p in cpoints)
            {
                var p1 = p.Subtract(0.1, 0.1);
                var p2 = p.Add(0.1, 0.1);
                page.DrawRectangle(p1.X, p1.Y, p2.X, p2.Y);
            }
        }
    }
}