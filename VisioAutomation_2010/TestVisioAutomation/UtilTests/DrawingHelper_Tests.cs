using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VA=VisioAutomation;
using VisioAutomation.Extensions;

namespace TestVisioAutomation
{
    [TestClass]
    public class DrawingHelper_Tests : VisioAutomationTest
    {
        [TestMethod]
        public void MultiStopGradient2()
        {
            var page1 = GetNewPage();

            // get a rect that is the size of the map
            var size = page1.GetSize();

            // create the gradient
            var gradient = new VA.Effects.MultiStopGradient();
            gradient.Add(new VA.Drawing.ColorRGB(0xFFF468), 0.3, 0.0);
            gradient.Add(new VA.Drawing.ColorRGB(0xFFF799), 1.0, 1.0);
            gradient.Direction = VA.Effects.MultiStopGradientDirection.LeftToRight;

            // the rext to fill with this gradient
            var rect = new VA.Drawing.Rectangle(0, 0, size.Width, size.Height);

            // along which axis to draw the gradient

            var shape = gradient.Draw(page1, rect);
            var shapes = VA.ShapeHelper.GetNestedShapes(shape);

            Assert.AreEqual(2, shapes.Count);

            var shapeids = shapes.Select(s => s.ID).ToArray();
            var formats = VA.Format.FormatHelper.GetShapeFormat(page1, shapeids);
        }

        [TestMethod]
        public void MultiStopGradient5()
        {
            var page1 = GetNewPage();

            // get a rect that is the size of the map
            page1.SetSize(10,10);
            var size = page1.GetSize();

            // create the gradient
            var gradient = new VA.Effects.MultiStopGradient();
            gradient.Add(new VA.Drawing.ColorRGB(0xFFF468), 0.3, 0.0);
            gradient.Add(new VA.Drawing.ColorRGB(0xFFF799), 1.0, 0.25);
            gradient.Add(new VA.Drawing.ColorRGB(0xFFC20E), 0.5, 0.50);
            gradient.Add(new VA.Drawing.ColorRGB(0xEB6119), 0.2, 0.90);
            gradient.Add(new VA.Drawing.ColorRGB(0xFBAF5D), 0.6, 1.0);
            gradient.Direction = VA.Effects.MultiStopGradientDirection.LeftToRight;


            // the rext to fill with this gradient
            var rect = new VA.Drawing.Rectangle(0, 0, size.Width, size.Height);

            // along which axis to draw the gradient

            var shape = gradient.Draw(page1, rect);
            var shapes = VA.ShapeHelper.GetNestedShapes(shape);

            Assert.AreEqual(5, shapes.Count);

            var shapeids = shapes.Select(s=>s.ID).ToArray();
            var formats = VA.Format.FormatHelper.GetShapeFormat(page1, shapeids);
        }
    }
}