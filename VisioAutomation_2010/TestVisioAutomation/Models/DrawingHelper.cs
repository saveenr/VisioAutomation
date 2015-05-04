using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Drawing;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace TestVisioAutomation
{
    [TestClass]
    public class DrawingHelper : VisioAutomationTest
    {
        [TestMethod]
        public void Drawing_CreateBoundingBox()
        {
            var doubles = new[] {0.0, 0.0, 1.0, -2.0};
            var points = Point.FromDoubles(doubles);
            var bb0 = new BoundingBox(points);
            var bb = bb0.Rectangle;
            Assert.AreEqual(0, bb.Left);
            Assert.AreEqual(0, bb.Top);
            Assert.AreEqual(1, bb.Right);
            Assert.AreEqual(-2, bb.Bottom);
        }
    }
}