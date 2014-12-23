using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using VisioAutomation.Extensions;

namespace TestVisioAutomation
{
    [TestClass]
    public class PieSliceTests : VisioAutomationTest
    {
        [TestMethod]
        public void Radial_DrawPieSlices()
        {
            var doc = this.GetNewDoc();
            var app = doc.Application;
            var page = app.ActivePage;

            var center = new VA.Drawing.Point(4, 5);
            double radius = 1.0;
            var values = new[] {1.0, 2.0};
            var slices = VA.Models.Charting.PieSlice.GetSlicesFromValues(center, radius, values);

            var shapes = new IVisio.Shape[values.Length];
            for (int i=0 ;i<values.Length;i++)
            {
                var slice = slices[i];
                var shape = slice.Render(page);
                shapes[i] = shape;
                shape.Text = values[i].ToString();
            }

            var shapeids = shapes.Select(s => s.ID).ToList();
            var xfrms = VA.Shapes.XFormCells.GetCells(page, shapeids);

            Assert.AreEqual("4.25 in", xfrms[0].PinX.Formula);
            Assert.AreEqual("5.5 in", xfrms[0].PinY.Formula);
            Assert.AreEqual("4 in", xfrms[1].PinX.Formula);
            Assert.AreEqual("4.9330127018922 in", xfrms[1].PinY.Formula);

            doc.Close(true);
        }
    }
}