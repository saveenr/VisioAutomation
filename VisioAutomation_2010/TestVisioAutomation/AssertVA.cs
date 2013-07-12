using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Drawing;
using VA=VisioAutomation;

namespace TestVisioAutomation
{
    public static class AssertVA
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

        public static void AreEqual(double x, double y, Size p, double delta)
        {
            Assert.AreEqual(x, p.Width, delta);
            Assert.AreEqual(y, p.Height, delta);
        }

        public static void AssertSnap(double ex, double ey, VA.Drawing.SnappingGrid g1, double ix, double iy, double delta)
        {
            AssertVA.AreEqual(ex, ey, g1.Snap(ix, iy), delta);
        }

        public static void AreEqual<T>(string formula, T result, VA.ShapeSheet.CellData<T> cd)
        {
            Assert.AreEqual(formula, cd.Formula);
            Assert.AreEqual(result, cd.Result);
        }

        public static void AreEqual<T>(string formula, T result, string af, T ar)
        {
            Assert.AreEqual(formula, af);
            Assert.AreEqual(result, ar);
        }


    }
}