using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;
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
            var points = VA.Drawing.Point.FromDoubles(doubles);
            var bb0 = new VA.Drawing.BoundingBox(points);
            var bb = bb0.Rectangle;
            Assert.AreEqual(0, bb.Left);
            Assert.AreEqual(0, bb.Top);
            Assert.AreEqual(1, bb.Right);
            Assert.AreEqual(-2, bb.Bottom);
        }
    }
}