using Microsoft.VisualStudio.TestTools.UnitTesting;
using VA = VisioAutomation;

namespace VisioAutomation_Tests.Models.Geometry
{
    [TestClass]
    public class DrawingHelper : VisioAutomationTest
    {
        [TestMethod]
        public void Drawing_CreateBoundingBox_0Points()
        {
            var doubles = new double[] { };
            var points = VA.Geometry.Point.FromDoubles(doubles);
            var bb = VisioAutomation.Models.Geometry.BoundingBoxBuilder.FromPoints(points);

            Assert.IsFalse(bb.HasValue);
        }

        [TestMethod]
        public void Drawing_CreateBoundingBox_1Point()
        {
            var doubles = new[] { 1.0, -2.0 };
            var points = VA.Geometry.Point.FromDoubles(doubles);
            var bb = VisioAutomation.Models.Geometry.BoundingBoxBuilder.FromPoints(points);

            Assert.IsTrue(bb.HasValue);
            Assert.AreEqual(1.0, bb.Value.Left);
            Assert.AreEqual(-2.0, bb.Value.Top);
            Assert.AreEqual(1.0, bb.Value.Right);
            Assert.AreEqual(-2.0, bb.Value.Bottom);
        }

        [TestMethod]
        public void Drawing_CreateBoundingBox_4Points()
        {
            var doubles = new[] {0.0, 0.0, 1.0, -2.0};
            var points = VA.Geometry.Point.FromDoubles(doubles);
            var bb = VisioAutomation.Models.Geometry.BoundingBoxBuilder.FromPoints(points);

            Assert.IsTrue(bb.HasValue);
            Assert.AreEqual(0, bb.Value.Left);
            Assert.AreEqual(0, bb.Value.Top);
            Assert.AreEqual(1, bb.Value.Right);
            Assert.AreEqual(-2, bb.Value.Bottom);
        }
    }
}