using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.DOM;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

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
            var slices = VA.Layout.Models.Radial.PieSlice.GetSlicesFromValues(center, radius, values);

            for (int i=0 ;i<values.Length;i++)
            {
                var slice = slices[i];
                var shape = slice.Render(page);
                shape.Text = values[i].ToString();

            }
        }
    }
}