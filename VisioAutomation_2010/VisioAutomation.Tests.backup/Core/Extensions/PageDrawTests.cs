using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;
using IVisio= Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;

namespace VisioAutomation_Tests.Core.Extensions
{
    [TestClass]
    public class PageDrawTests : VisioAutomationTest
    {
        [TestMethod]
        public void Page_Draw_Line()
        {
            var page1 = this.GetNewPage();
            var p0 = new VA.Geometry.Point(0, 0);
            var p1 = new VA.Geometry.Point(3, 2);
            var s0 = page1.DrawLine(p0, p1);
            page1.Delete(0);
        }

        [TestMethod]
        public void Page_Draw_Spline()
        {
            var page1 = this.GetNewPage();
            var points = new[]
            {
                new VA.Geometry.Point(0, 0),
                new VA.Geometry.Point(3, 3),
                new VA.Geometry.Point(2, 0)
            };

            var doubles_array = VA.Geometry.Point.ToDoubles(points).ToArray();
            var s0 = page1.DrawSpline(doubles_array, 0, 0);

            page1.Delete(0);
        }

        [TestMethod]
        public void Page_Draw_RoundedRectangle()
        {
            var page1 = this.GetNewPage();
            var rect = new VA.Geometry.Rectangle(1, 1, 3, 2);
            // draw an inital framing rectangle so the coordinates are easy to calculate
            var s0 = page1.DrawRectangle(rect);
            double width = rect.Width;
            double height = rect.Height;
            double delta = 1.0 / 8.0;

            var o = new VA.Geometry.Point(0, 0);

            var a = new VA.Geometry.Point(o.X + delta, o.Y);
            var b = new VA.Geometry.Point(o.X, o.Y + delta);
            var c = new VA.Geometry.Point(o.X, o.Y + height - delta);
            var d = new VA.Geometry.Point(o.X + delta, o.Y + height);
            var e = new VA.Geometry.Point(o.X + width - delta, o.Y + height);
            var f = new VA.Geometry.Point(o.X + width, o.Y + height - delta);
            var g = new VA.Geometry.Point(o.X + width, o.Y + delta);
            var h = new VA.Geometry.Point(o.X + width - delta, o.Y);

            var bottom_left_curve = s0.DrawQuarterArc(a, b, IVisio.VisArcSweepFlags.visArcSweepFlagConcave);
            var left_side = s0.DrawLine(b, c);
            var top_left_curve = s0.DrawQuarterArc(c, d, IVisio.VisArcSweepFlags.visArcSweepFlagConvex);
            var top_side = s0.DrawLine(d, e);
            var top_right_curve = s0.DrawQuarterArc(e, f, IVisio.VisArcSweepFlags.visArcSweepFlagConcave);
            var right_side = s0.DrawLine(f, g);
            var bottom_right_curve = s0.DrawQuarterArc(g, h, IVisio.VisArcSweepFlags.visArcSweepFlagConvex);
            var bottom_side = s0.DrawLine(h, a);

            // delete the framing rectangle
            s0.DeleteSection((short) IVisio.VisSectionIndices.visSectionFirstComponent);

            page1.Delete(0);
        }
    }
}
