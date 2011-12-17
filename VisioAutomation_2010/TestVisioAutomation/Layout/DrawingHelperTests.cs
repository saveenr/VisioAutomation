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
        public void BoundingBox()
        {
            var bb0 = new VA.Drawing.BoundingBox(VA.Drawing.Point.FromDoubles(new[] {0.0, 0.0, 1.0, -2.0}));
            var bb = bb0.Rectangle;
            Assert.AreEqual(0, bb.Left);
            Assert.AreEqual(0, bb.Top);
            Assert.AreEqual(1, bb.Right);
            Assert.AreEqual(-2, bb.Bottom);
        }
    }
}