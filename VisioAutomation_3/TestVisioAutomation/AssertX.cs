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

        public static void AreEqual(double x, double y, Point p2, double delta)
        {
            Assert.AreEqual(x, p2.X, delta);
            Assert.AreEqual(y, p2.Y, delta);
        }
    }
}