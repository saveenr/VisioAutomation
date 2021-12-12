using VisioAutomation.Extensions;

namespace VSamples.Samples.Misc
{
    public  class BezierEllipse : SampleMethodBase
    {
        public override void RunSample()
        {
            // Draw an approximation of an ellipse using Bezier Curves

            var page = SampleEnvironment.Application.ActiveDocument.Pages.Add();

            var curve0 = VisioAutomation.Models.Geometry.BezierCurve.FromEllipse(
                new VisioAutomation.Core.Point(2, 4),
                new VisioAutomation.Core.Size(1, 0.5));
            var s0 = page.DrawBezier(curve0.ControlPoints);
            s0.Text = "Bezier approximating ellipse";
        }
    }
}