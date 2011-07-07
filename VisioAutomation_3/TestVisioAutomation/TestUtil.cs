using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Drawing;

namespace TestVisioAutomation
{
    public static class TestUtil
    {
        public static void AreEqual(double x, double y, Point p, double delta)
        {
            Assert.AreEqual(x,p.X,delta);
            Assert.AreEqual(y,p.Y, delta);
        }

        public static void AreEqual(double x, double y, Size p, double delta)
        {
            Assert.AreEqual(x, p.Width, delta);
            Assert.AreEqual(y, p.Height, delta);
        }

    }
}