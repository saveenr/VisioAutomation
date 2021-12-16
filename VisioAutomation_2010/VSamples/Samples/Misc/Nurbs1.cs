using System.Linq;
using VisioAutomation.Extensions;

namespace VSamples.Samples.Misc
{
    public class Nurbs1 : SampleMethodBase
    {
        public override void Run()
        {
            // Draw a simple NURBS
            // Example from this page:http://www.robthebloke.org/opengl_programming.html

            var page = SampleEnvironment.Application.ActiveDocument.Pages.Add();

            var points = new[]
            {
                new VisioAutomation.Core.Point(10, 10),
                new VisioAutomation.Core.Point(5, 10),
                new VisioAutomation.Core.Point(-5, 5),
                new VisioAutomation.Core.Point(-10, 5),
                new VisioAutomation.Core.Point(-4, 10),
                new VisioAutomation.Core.Point(-4, 5),
                new VisioAutomation.Core.Point(-8, 1)
            };

            var origin = new VisioAutomation.Core.Point(4, 4);
            var scale = new VisioAutomation.Core.Size(1.0 / 4.0, 1.0 / 4.0);

            var controlpoints = points.Select(x => (x * scale) + origin).ToList();
            var knots = new double[] {0, 0, 0, 0, 1, 2, 3, 4, 4, 4, 4};
            var degree = 3;
            var weights = controlpoints.Select(i => 1.0).ToList();

            var s0 = page.DrawNurbs(controlpoints, knots, weights, degree);
            s0.Text = "Generic NURBS shape";
        }
    }
}