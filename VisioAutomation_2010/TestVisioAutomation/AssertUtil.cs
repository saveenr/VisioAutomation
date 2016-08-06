using Microsoft.VisualStudio.TestTools.UnitTesting;
using VADRAW = VisioAutomation.Drawing;
using VASS = VisioAutomation.ShapeSheet;

namespace TestVisioAutomation
{
    public static class AssertUtil
    {
        public static void AreEqual(VADRAW.Point point, VADRAW.Point actual_point, double delta)
        {
            Assert.AreEqual(point.X, actual_point.X,delta);
            Assert.AreEqual(point.Y, actual_point.Y, delta);
        }

        public static void AreEqual(double x, double y, VADRAW.Point actual_point, double delta)
        {
            Assert.AreEqual(x, actual_point.X, delta);
            Assert.AreEqual(y, actual_point.Y, delta);
        }

        public static void AreEqual(double left, double bottom, double right, double top, VADRAW.Rectangle actual_rect, double delta)
        {
            Assert.AreEqual(left, actual_rect.Left, delta);
            Assert.AreEqual(bottom, actual_rect.Bottom, delta);
            Assert.AreEqual(right, actual_rect.Right, delta);
            Assert.AreEqual(top, actual_rect.Top, delta);
        }

        public static void AreEqual(double width, double height, VADRAW.Size actual_size, double delta)
        {
            Assert.AreEqual(width, actual_size.Width, delta);
            Assert.AreEqual(height, actual_size.Height, delta);
        }

        public static void AssertSnap(double ex, double ey, VADRAW.SnappingGrid g1, double ix, double iy, double delta)
        {
            AssertUtil.AreEqual(ex, ey, g1.Snap(ix, iy), delta);
        }

        public static void AreEqual<T>(string formula, T result, VASS.CellData<T> actual_celldata)
        {
            Assert.AreEqual(formula, actual_celldata.Formula);
            Assert.AreEqual(result, actual_celldata.Result);
        }

        public static void AreEqual<T>(string formula, T result, string actual_formula, T actual_result)
        {
            Assert.AreEqual(formula, actual_formula);
            Assert.AreEqual(result, actual_result);
        }
    }
}