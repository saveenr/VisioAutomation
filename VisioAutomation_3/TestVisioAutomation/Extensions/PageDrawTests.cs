using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation;
using VisioAutomation.Extensions;
using System.Linq;
using IVisio= Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;

namespace TestVisioAutomation
{
    [TestClass]
    public class PageDrawTests : VisioAutomationTest
    {
        [TestMethod]
        public void DrawLine1()
        {
            var page1 = GetNewPage();
            var s0 = page1.DrawLine(new VA.Drawing.Point(0, 0), new VA.Drawing.Point(3, 2));
            page1.Delete(0);
        }

        [TestMethod]
        public void DrawLine()
        {
            var page1 = GetNewPage();
            var s0 = page1.DrawLine(new VA.Drawing.Point(0, 0), new VA.Drawing.Point(3, 3));
            page1.Delete(0);
        }

        [TestMethod]
        public void DrawSpline()
        {
            var page1 = GetNewPage();

            var points = new[]
                             {
                                 new VA.Drawing.Point(0, 0), 
                                 new VA.Drawing.Point(3, 3),
                                 new VA.Drawing.Point(2, 0)
                             };

            var doubles_array = VA.Drawing.DrawingUtil.PointsToDoubles(points).ToArray();
            var s0 = page1.DrawSpline(doubles_array, 0, 0);

            page1.Delete(0);
        }

        [TestMethod]
        public void DrawRoundedRectReometry()
        {
            var page1 = GetNewPage();
            var rect = new VA.Drawing.Rectangle(1, 1, 3, 2);
            // draw an inital framing rectangle so the coordinates are easy to calculate
            var s0 = page1.DrawRectangle(rect);
            double width = rect.Width;
            double height = rect.Height;
            double delta = 1.0/8.0;

            var o = new VA.Drawing.Point(0, 0);

            var a = new VA.Drawing.Point(o.X + delta, o.Y);
            var b = new VA.Drawing.Point(o.X, o.Y + delta);
            var c = new VA.Drawing.Point(o.X, o.Y + height - delta);
            var d = new VA.Drawing.Point(o.X + delta, o.Y + height);
            var e = new VA.Drawing.Point(o.X + width - delta, o.Y + height);
            var f = new VA.Drawing.Point(o.X + width, o.Y + height - delta);
            var g = new VA.Drawing.Point(o.X + width, o.Y + delta);
            var h = new VA.Drawing.Point(o.X + width - delta, o.Y);

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

        [TestMethod]
        public void DropManyShapes()
        {
            var page1 = GetNewPage();            
            var stencil = "basic_u.vss";

            short flags = (short)IVisio.VisOpenSaveArgs.visOpenRO | (short)IVisio.VisOpenSaveArgs.visOpenDocked;
            var app = page1.Application;
            var documents = app.Documents;
            var stencil_doc = documents.OpenEx(stencil, flags);

            var masters1 = stencil_doc.Masters;
            var masters = new [] {masters1["Rounded Rectangle"], masters1["Ellipse"]};
            var points = new [] {new VA.Drawing.Point(1, 2), new VA.Drawing.Point(3, 4)};
            Assert.AreEqual(0, page1.Shapes.Count);
            var shapeids = page1.DropManyU(masters, points);
            Assert.AreEqual(2, page1.Shapes.Count);
            Assert.AreEqual(2, shapeids.Length );

            var s0 = page1.Shapes[shapeids[0]];
            var s1 = page1.Shapes[shapeids[1]];

            Assert.AreEqual( masters[0].NameU, s0.Master.NameU );
            Assert.AreEqual(masters[1].NameU, s1.Master.NameU);
            
            page1.Delete(0);
        }
    }
}