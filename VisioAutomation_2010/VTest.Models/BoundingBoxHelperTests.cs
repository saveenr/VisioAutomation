using MUT=Microsoft.VisualStudio.TestTools.UnitTesting;
using VA = VisioAutomation;

namespace VTest.Models
{
    [MUT.TestClass]
    public class BoundingBoxHelperTests : Framework.VTest
    {
        [MUT.TestMethod]
        public void Drawing_CreateBoundingBox_0Points()
        {
            var doubles = new double[] { };
            var points = VA.Core.Point.FromDoubles(doubles);
            var bb = VisioAutomation.Models.Geometry.BoundingBoxBuilder.FromPoints(points);

            MUT.Assert.IsFalse(bb.HasValue);
        }

        [MUT.TestMethod]
        public void Drawing_CreateBoundingBox_1Point()
        {
            var doubles = new[] { 1.0, -2.0 };
            var points = VA.Core.Point.FromDoubles(doubles);
            var bb = VisioAutomation.Models.Geometry.BoundingBoxBuilder.FromPoints(points);

            MUT.Assert.IsTrue(bb.HasValue);
            MUT.Assert.AreEqual(1.0, bb.Value.Left);
            MUT.Assert.AreEqual(-2.0, bb.Value.Top);
            MUT.Assert.AreEqual(1.0, bb.Value.Right);
            MUT.Assert.AreEqual(-2.0, bb.Value.Bottom);
        }

        [MUT.TestMethod]
        public void Drawing_CreateBoundingBox_4Points()
        {
            var doubles = new[] {0.0, 0.0, 1.0, -2.0};
            var points = VA.Core.Point.FromDoubles(doubles);
            var bb = VisioAutomation.Models.Geometry.BoundingBoxBuilder.FromPoints(points);

            MUT.Assert.IsTrue(bb.HasValue);
            MUT.Assert.AreEqual(0, bb.Value.Left);
            MUT.Assert.AreEqual(0, bb.Value.Top);
            MUT.Assert.AreEqual(1, bb.Value.Right);
            MUT.Assert.AreEqual(-2, bb.Value.Bottom);
        }
    }
}