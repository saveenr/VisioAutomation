using System.Globalization;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation_Tests.Models
{
    [TestClass]
    public class PieSliceTests : VisioAutomationTest
    {
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
        public void Radial_DrawPieSlices()
        {
            var doc = this.GetNewDoc();
            var app = doc.Application;
            var page = app.ActivePage;

            var center = new VA.Drawing.Point(4, 5);
            double radius = 1.0;
            var values = new[] {1.0, 2.0};
            var slices = VisioAutomation.Models.Charting.PieSlice.GetSlicesFromValues(center, radius, values);

            var shapes = new IVisio.Shape[values.Length];
            for (int i=0 ;i<values.Length;i++)
            {
                var slice = slices[i];
                var shape = slice.Render(page);
                shapes[i] = shape;
                shape.Text = values[i].ToString(CultureInfo.InvariantCulture);
            }

            var shapeids = shapes.Select(s => s.ID).ToList();
            var xfrms = VisioAutomation.Shapes.ShapeXFormCells.GetCells(page, shapeids);

            Assert.AreEqual("4.25 in", xfrms[0].PinX.Formula);
            Assert.AreEqual("5.5 in", xfrms[0].PinY.Formula);
            Assert.AreEqual("4 in", xfrms[1].PinX.Formula);
            Assert.AreEqual("4.9330127018922 in", xfrms[1].PinY.Formula);

            doc.Close(true);
        }
    }
}