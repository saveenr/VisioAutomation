using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace TestVisioAutomation
{
    [TestClass]
    public class BezierTests : VisioAutomationTest
    {
        private double delta = 0.00000000001;
       double pi2 =       System.Math.PI * 2;
       double pihalf =    System.Math.PI / 2;
       double piquarter = System.Math.PI / 4;
        [TestMethod]
        public void TestBezierFromArcs()
        {
            // 0 width slice - 0 degrees
            var s1 = VA.Drawing.BezierSegment.FromArc(0.0, 0.0 );
            Assert.AreEqual(1, s1.Count());
            AssertX.AreEqual(s1[0].Start, s1[s1.Length - 1].End, delta);

            // 0 width slice - 45 degrees
            var s1x = VA.Drawing.BezierSegment.FromArc(piquarter, piquarter);
            Assert.AreEqual(1, s1x.Count());
            AssertX.AreEqual(s1x[0].Start, s1x[s1.Length - 1].End, delta);

            // a circle
            var s2 = VA.Drawing.BezierSegment.FromArc(0.0, pi2);
            Assert.AreEqual(4,s2.Count());
            AssertX.AreEqual(s2[0].Start, s2[s2.Length - 1].End,delta);

            // angles within first quadrant
            var s3 = VA.Drawing.BezierSegment.FromArc(piquarter-0.1, piquarter+0.2);
            Assert.AreEqual(1, s3.Count());

            // angles from first to 2nd quadrant
            var s4 = VA.Drawing.BezierSegment.FromArc(piquarter - 0.1, pihalf + piquarter) ;
            Assert.AreEqual(2, s4.Count());

            // half circle - top
            var s5 = VA.Drawing.BezierSegment.FromArc(0.0, System.Math.PI);
            Assert.AreEqual(2, s5.Count());

            // half circle - bottom
            var s6 = VA.Drawing.BezierSegment.FromArc(System.Math.PI, pi2);
            Assert.AreEqual(2, s6.Count());

            // half circle - bottom
            var s7 = VA.Drawing.BezierSegment.FromArc(pihalf, System.Math.PI + pihalf);
            Assert.AreEqual(2, s7.Count());

            // partial all quadrants
            var s8 = VA.Drawing.BezierSegment.FromArc(piquarter, pi2 - piquarter);
            Assert.AreEqual(4, s8.Count());

            // full circle
            var s9 = VA.Drawing.BezierSegment.FromArc(piquarter, pi2*10 + piquarter);
            Assert.AreEqual(4, s8.Count());
        }
    }
}