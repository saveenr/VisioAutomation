using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VisioAutomation_Tests.Models.Geometry
{
    [TestClass]
    public class BezierTests : VisioAutomationTest
    {
        private double delta = 0.00000000001;
        private double pi2 = Math.PI*2;
        private double pihalf = Math.PI/2;
        private double piquarter = Math.PI/4;

        [TestMethod]
        public void TestBezierFromArcs()
        {
            // 0 width slice - 0 degrees
            var s1 = VisioAutomation.Models.Geometry.BezierSegment.FromArc(0.0, 0.0);
            Assert.AreEqual(1, s1.Length);
            Assert.AreEqual(s1[0].Start.X, s1[s1.Length - 1].End.X, this.delta);
            Assert.AreEqual(s1[0].Start.Y, s1[s1.Length - 1].End.Y, this.delta);

            // 0 width slice - 45 degrees
            var s1_x = VisioAutomation.Models.Geometry.BezierSegment.FromArc(this.piquarter, this.piquarter);
            Assert.AreEqual(1, s1_x.Length);
            Assert.AreEqual(s1_x[0].Start.X, s1_x[s1.Length - 1].End.X, this.delta);
            Assert.AreEqual(s1_x[0].Start.Y, s1_x[s1.Length - 1].End.Y, this.delta);

            // a circle
            var s2 = VisioAutomation.Models.Geometry.BezierSegment.FromArc(0.0, this.pi2);
            Assert.AreEqual(4, s2.Length);
            Assert.AreEqual(s2[0].Start.X, s2[s2.Length - 1].End.X, this.delta);
            Assert.AreEqual(s2[0].Start.Y, s2[s2.Length - 1].End.Y, this.delta);

            // angles within first quadrant
            var s3 = VisioAutomation.Models.Geometry.BezierSegment.FromArc(this.piquarter - 0.1, this.piquarter + 0.2);
            Assert.AreEqual(1, s3.Length);

            // angles from first to 2nd quadrant
            var s4 = VisioAutomation.Models.Geometry.BezierSegment.FromArc(this.piquarter - 0.1, this.pihalf + this.piquarter);
            Assert.AreEqual(2, s4.Length);

            // half circle - top
            var s5 = VisioAutomation.Models.Geometry.BezierSegment.FromArc(0.0, Math.PI);
            Assert.AreEqual(2, s5.Length);

            // half circle - bottom
            var s6 = VisioAutomation.Models.Geometry.BezierSegment.FromArc(Math.PI, this.pi2);
            Assert.AreEqual(2, s6.Length);

            // half circle - bottom
            var s7 = VisioAutomation.Models.Geometry.BezierSegment.FromArc(this.pihalf, Math.PI + this.pihalf);
            Assert.AreEqual(2, s7.Length);

            // partial all quadrants
            var s8 = VisioAutomation.Models.Geometry.BezierSegment.FromArc(this.piquarter, this.pi2 - this.piquarter);
            Assert.AreEqual(4, s8.Length);

            // full circle
            var s9 = VisioAutomation.Models.Geometry.BezierSegment.FromArc(this.piquarter, this.pi2*10 + this.piquarter);
            Assert.AreEqual(4, s8.Length);
        }
    }
}