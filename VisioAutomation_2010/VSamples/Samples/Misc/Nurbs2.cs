using System.Linq;
using VisioAutomation.Extensions;

namespace VSamples.Samples.Misc
{
    public class Nurbs2 : SampleMethodBase
    {
        public override void RunSample()
        {
            // Draw a simple NURBS
            // Example from Graham Wideman's book

            var page = SampleEnvironment.Application.ActiveDocument.Pages.Add();

            var points = new[]
            {
                new VisioAutomation.Core.Point(0.2500, 0.2500),
                new VisioAutomation.Core.Point(0.2500, 0.7500),
                new VisioAutomation.Core.Point(0.4063, 0.8125),
                new VisioAutomation.Core.Point(0.5625, 0.3750),
                new VisioAutomation.Core.Point(0.5538, 0.8125),
                new VisioAutomation.Core.Point(0.7600, 0.7500),
                new VisioAutomation.Core.Point(0.7600, 0.2500)
            };

            var origin = new VisioAutomation.Core.Point(4, 4);
            var scale = new VisioAutomation.Core.Size(4, 4);

            var controlpoints = points.Select(x => (x * scale) + origin).ToList();
            var knots = new double[] {0, 0, 0, 0, 25, 50, 75, 100, 100, 100, 100};
            var degree = 3;
            var Weights = controlpoints.Select(i => 1.0).ToList();

            var s0 = page.DrawNurbs(controlpoints, knots, Weights, degree);
            s0.Text = "Generic NURBS shape";
        }
    }
}