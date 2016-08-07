using System.Globalization;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;
using VACHART=VisioAutomation.Models.Charting;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation_Tests.Models
{
    [TestClass]
    public class Chart_Tests : VisioAutomationTest
    {
        [TestMethod]
        public void Chart_Bar1()
        {
            var doc = this.GetNewDoc();
            var app = doc.Application;
            var page = app.ActivePage;

            var center = new VA.Drawing.Point(4, 5);
            double radius = 1.0;
            var values = new[] {1.0, 2.0};
            var slices = VACHART.PieSlice.GetSlicesFromValues(center, radius, values);

            var shapes = new IVisio.Shape[values.Length];
            for (int i=0 ;i<values.Length;i++)
            {
                var slice = slices[i];
                var shape = slice.Render(page);
                shapes[i] = shape;
                shape.Text = values[i].ToString(CultureInfo.InvariantCulture);
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