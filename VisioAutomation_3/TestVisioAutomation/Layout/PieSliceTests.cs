using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace TestVisioAutomation
{
    [TestClass]
    public class PieSliceTests : VisioAutomationTest
    {
        [TestMethod]
        public void DrawSliceRanges()
        {
            var app = this.GetVisioApplication();
            var doc = this.GetNewDoc();
            var page = app.ActivePage;

            int n = 36;
            double start_angle = 0.0;
            double radius = 1.0;
            double cx = 0.0;
            double cy = 2.0;
            double angle_step = System.Math.PI*2.0/ (n - 1);

            foreach (double end_angle in Enumerable.Range(0, n).Select(i => i * angle_step))
            {
                var center = new VA.Drawing.Point(cx, cy);
                VA.Layout.DrawingtHelper.DrawPieSlice(page, center, radius, start_angle, end_angle);
                cx += 2.5;
            }

            page.ResizeToFitContents(1,1);
            doc.Close(true);
        }

        [TestMethod]
        public void DrawThickArcRanges()
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
                VA.Layout.DrawingtHelper.DrawArc(page, center, radius - 0.2, radius, start_angle, end_angle);
                cx += 2.5;
            }

            page.ResizeToFitContents(1, 1);
            doc.Close(true);
        }

    }
}