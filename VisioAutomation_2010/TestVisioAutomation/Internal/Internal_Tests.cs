using System;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VA = VisioAutomation;

namespace TestVisioAutomation.Internal
{
    [TestClass]
    public class Internal_Tests
    {
        [TestMethod]
        public void Internal_ValidateSnappingGrid()
        {
            double delta = 0.000000001;

            var g1 = new VA.Drawing.SnappingGrid(1.0, 1.0);

            AssertVA.AssertSnap(0.0, 0.0, g1, 0.0, 0.0, delta);
            AssertVA.AssertSnap(0.0, 0.0, g1, 0.3, 0.3, delta);
            AssertVA.AssertSnap(0.0, 0.0, g1, 0.49999, 0.49999, delta);
            AssertVA.AssertSnap(1.0, 1.0, g1, 0.5, 0.5, delta);
            AssertVA.AssertSnap(1.0, 1.0, g1, 0.500001, 0.500001, delta);
            AssertVA.AssertSnap(1.0, 1.0, g1, 1.0, 1.0, delta);
            AssertVA.AssertSnap(1.0, 1.0, g1, 1.3, 1.3, delta);
            AssertVA.AssertSnap(1.0, 1.0, g1, 1.49999, 1.49999, delta);
            AssertVA.AssertSnap(2.0, 2.0, g1, 1.5, 1.5, delta);
            AssertVA.AssertSnap(2.0, 2.0, g1, 1.500001, 1.500001, delta);

            var g2 = new VA.Drawing.SnappingGrid(1.0, 0.3);

            AssertVA.AssertSnap(0.0, 0.0, g2, 0.0, 0.0, delta);
            AssertVA.AssertSnap(0.0, 0.0, g2, 0.3, 0.1, delta);
            AssertVA.AssertSnap(0.0, 0.0, g2, 0.49999, 0.149, delta);
            AssertVA.AssertSnap(1.0, 0.3, g2, 0.5, 0.3, delta);
            AssertVA.AssertSnap(1.0, 0.3, g2, 0.500001, 0.30001, delta);
        }

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
                var s1 = VA.Drawing.BezierSegment.FromArc(0.0, 0.0);
                Assert.AreEqual(1, s1.Count());
                AssertVA.AreEqual(s1[0].Start, s1[s1.Length - 1].End, this.delta);

                // 0 width slice - 45 degrees
                var s1x = VA.Drawing.BezierSegment.FromArc(this.piquarter, this.piquarter);
                Assert.AreEqual(1, s1x.Count());
                AssertVA.AreEqual(s1x[0].Start, s1x[s1.Length - 1].End, this.delta);

                // a circle
                var s2 = VA.Drawing.BezierSegment.FromArc(0.0, this.pi2);
                Assert.AreEqual(4, s2.Count());
                AssertVA.AreEqual(s2[0].Start, s2[s2.Length - 1].End, this.delta);

                // angles within first quadrant
                var s3 = VA.Drawing.BezierSegment.FromArc(this.piquarter - 0.1, this.piquarter + 0.2);
                Assert.AreEqual(1, s3.Count());

                // angles from first to 2nd quadrant
                var s4 = VA.Drawing.BezierSegment.FromArc(this.piquarter - 0.1, this.pihalf + this.piquarter);
                Assert.AreEqual(2, s4.Count());

                // half circle - top
                var s5 = VA.Drawing.BezierSegment.FromArc(0.0, Math.PI);
                Assert.AreEqual(2, s5.Count());

                // half circle - bottom
                var s6 = VA.Drawing.BezierSegment.FromArc(Math.PI, this.pi2);
                Assert.AreEqual(2, s6.Count());

                // half circle - bottom
                var s7 = VA.Drawing.BezierSegment.FromArc(this.pihalf, Math.PI + this.pihalf);
                Assert.AreEqual(2, s7.Count());

                // partial all quadrants
                var s8 = VA.Drawing.BezierSegment.FromArc(this.piquarter, this.pi2 - this.piquarter);
                Assert.AreEqual(4, s8.Count());

                // full circle
                var s9 = VA.Drawing.BezierSegment.FromArc(this.piquarter, this.pi2*10 + this.piquarter);
                Assert.AreEqual(4, s8.Count());
            }
        }
    }
}