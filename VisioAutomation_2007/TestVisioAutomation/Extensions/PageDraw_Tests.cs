using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;
using System.Linq;
using IVisio= Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;

namespace TestVisioAutomation
{
    [TestClass]
    public class PageDraw_Tests : VisioAutomationTest
    {
        [TestMethod]
        public void Page_Draw_Line()
        {
            var page1 = GetNewPage();
            var p0 = new VA.Drawing.Point(0, 0);
            var p1 = new VA.Drawing.Point(3, 2);
            var s0 = page1.DrawLine(p0, p1);
            page1.Delete(0);
        }


        [TestMethod]
        public void Page_Draw_Spline()
        {
            var page1 = GetNewPage();
            var points = new[]
                             {
                                 new VA.Drawing.Point(0, 0), 
                                 new VA.Drawing.Point(3, 3),
                                 new VA.Drawing.Point(2, 0)
                             };

            var doubles_array = VA.Drawing.Point.ToDoubles(points).ToArray();
            var s0 = page1.DrawSpline(doubles_array, 0, 0);

            page1.Delete(0);
        }

        [TestMethod]
        public void Page_Draw_RoundedRectangle()
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
        public void Page_Draw_PieSlices()
        {
            var app = this.GetVisioApplication();
            var doc = this.GetNewDoc();
            var page = app.ActivePage;

            int n = 36;
            double start_angle = 0.0;
            double radius = 1.0;
            double cx = 0.0;
            double cy = 2.0;
            double angle_step = System.Math.PI * 2.0 / (n - 1);

            foreach (double end_angle in Enumerable.Range(0, n).Select(i => i * angle_step))
            {
                var center = new VA.Drawing.Point(cx, cy);
                var ps = new VA.Models.Charting.PieSlice(center, radius, start_angle, end_angle);
                ps.Render(page);
                cx += 2.5;
            }

            var bordersize = new VA.Drawing.Size(1, 1);
            page.ResizeToFitContents(bordersize);

            doc.Close(true);
        }

        [TestMethod]
        public void Page_Draw_DoughnutSlices()
        {
            var app = this.GetVisioApplication();
            var doc = this.GetNewDoc();
            var page = app.ActivePage;

            int n = 36;
            double start_angle = 0.0;
            double radius = 1.0;
            double cx = 0.0;
            double cy = 2.0;
            double angle_step = System.Math.PI * 2.0 / (n - 1);

            foreach (double end_angle in Enumerable.Range(0, n).Select(i => i * angle_step))
            {
                var center = new VA.Drawing.Point(cx, cy);
                var slice = new VA.Models.Charting.PieSlice(center, start_angle, end_angle, radius - 0.2, radius);
                slice.Render(page);
                cx += 2.5;
            }

            var bordersize = new VA.Drawing.Size(1, 1);
            page.ResizeToFitContents(bordersize);
            doc.Close(true);
        }

        [TestMethod]
        public void Page_Drop_ManyU()
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