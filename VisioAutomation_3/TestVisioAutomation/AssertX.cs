using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Drawing;

namespace TestVisioAutomation
{
    public static class AssertX
    {
        public static void AreEqual(Point p1, Point p2, double delta)
        {
            Assert.AreEqual(p1.X,p2.X,delta);
            Assert.AreEqual(p1.Y, p2.Y, delta);
        }

        public static void AreEqual(double x, double y, Point p, double delta)
        {
            Assert.AreEqual(x, p.X, delta);
            Assert.AreEqual(y, p.Y, delta);
        }

        public static void AreEqual(double left, double bottom, double right, double top, Rectangle r, double delta)
        {
            Assert.AreEqual(left, r.Left, delta);
            Assert.AreEqual(bottom, r.Bottom, delta);
            Assert.AreEqual(right, r.Right, delta);
            Assert.AreEqual(top, r.Top, delta);
        }
    }
}