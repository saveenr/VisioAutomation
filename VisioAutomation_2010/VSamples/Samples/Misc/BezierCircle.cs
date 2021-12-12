using VisioAutomation.Extensions;

namespace VSamples.Samples.Misc
{
    public class BezierCircle : SampleMethodBase
    {
        public override void RunSample()
        {
            // Draw an approximation of a circle using Bezier Curves

            var page = SampleEnvironment.Application.ActiveDocument.Pages.Add();

            var curve0 = VisioAutomation.Models.Geometry.BezierCurve.FromEllipse(
                new VisioAutomation.Core.Point(5, 4),
                new VisioAutomation.Core.Size(1, 1));

            var s0 = page.DrawBezier(curve0.ControlPoints);
            s0.Text = "Bezier approximating circle";
        }
    }
}