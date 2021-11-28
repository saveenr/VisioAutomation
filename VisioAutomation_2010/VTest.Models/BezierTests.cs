using MUT=Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VTest.Models.Geometry
{
    [MUT.TestClass]
    public class BezierTests : Framework.VTest
    {
        private double delta = 0.00000000001;
        private double pi2 = System.Math.PI*2;
        private double pihalf = System.Math.PI/2;
        private double piquarter = System.Math.PI/4;

        [MUT.TestMethod]
        public void TestBezierFromArcs()
        {
            // 0 width slice - 0 degrees
            var s1 = VisioAutomation.Models.Geometry.BezierSegment.FromArc(0.0, 0.0);
            MUT.Assert.AreEqual(1, s1.Length);
            MUT.Assert.AreEqual(s1[0].Start.X, s1[s1.Length - 1].End.X, this.delta);
            MUT.Assert.AreEqual(s1[0].Start.Y, s1[s1.Length - 1].End.Y, this.delta);

            // 0 width slice - 45 degrees
            var s1_x = VisioAutomation.Models.Geometry.BezierSegment.FromArc(this.piquarter, this.piquarter);
            MUT.Assert.AreEqual(1, s1_x.Length);
            MUT.Assert.AreEqual(s1_x[0].Start.X, s1_x[s1.Length - 1].End.X, this.delta);
            MUT.Assert.AreEqual(s1_x[0].Start.Y, s1_x[s1.Length - 1].End.Y, this.delta);

            // a circle
            var s2 = VisioAutomation.Models.Geometry.BezierSegment.FromArc(0.0, this.pi2);
            MUT.Assert.AreEqual(4, s2.Length);
            MUT.Assert.AreEqual(s2[0].Start.X, s2[s2.Length - 1].End.X, this.delta);
            MUT.Assert.AreEqual(s2[0].Start.Y, s2[s2.Length - 1].End.Y, this.delta);

            // angles within first quadrant
            var s3 = VisioAutomation.Models.Geometry.BezierSegment.FromArc(this.piquarter - 0.1, this.piquarter + 0.2);
            MUT.Assert.AreEqual(1, s3.Length);

            // angles from first to 2nd quadrant
            var s4 = VisioAutomation.Models.Geometry.BezierSegment.FromArc(this.piquarter - 0.1, this.pihalf + this.piquarter);
            MUT.Assert.AreEqual(2, s4.Length);

            // half circle - top
            var s5 = VisioAutomation.Models.Geometry.BezierSegment.FromArc(0.0, System.Math.PI);
            MUT.Assert.AreEqual(2, s5.Length);

            // half circle - bottom
            var s6 = VisioAutomation.Models.Geometry.BezierSegment.FromArc(System.Math.PI, this.pi2);
            MUT.Assert.AreEqual(2, s6.Length);

            // half circle - bottom
            var s7 = VisioAutomation.Models.Geometry.BezierSegment.FromArc(this.pihalf, System.Math.PI + this.pihalf);
            MUT.Assert.AreEqual(2, s7.Length);

            // partial all quadrants
            var s8 = VisioAutomation.Models.Geometry.BezierSegment.FromArc(this.piquarter, this.pi2 - this.piquarter);
            MUT.Assert.AreEqual(4, s8.Length);

            // full circle
            var s9 = VisioAutomation.Models.Geometry.BezierSegment.FromArc(this.piquarter, this.pi2*10 + this.piquarter);
            MUT.Assert.AreEqual(4, s8.Length);
        }
    }
}