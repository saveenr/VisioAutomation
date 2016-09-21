using Microsoft.VisualStudio.TestTools.UnitTesting;
using VA=VisioAutomation;
using VADRAW = VisioAutomation.Drawing;

namespace VisioAutomation_Tests
{
    public static class AssertUtil
    {
        public static void FileExists(string filename)
        {
            Assert.IsTrue(System.IO.File.Exists(filename));
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

        public static void AssertSnap(double expected_x, double expected_y, VA.Scripting.Utilities.SnappingGrid snapgrid, double input_x, double input_y, double delta)
        {
            var snapped = snapgrid.Snap(input_x, input_y);
            Assert.AreEqual(expected_x, snapped.X, delta);
            Assert.AreEqual(expected_y, snapped.Y, delta);
        }

        public static void AreEqual<TResult>(string formula, TResult result, string actual_formula, TResult actual_result)
        {
            Assert.AreEqual(formula, actual_formula);
            Assert.AreEqual(result, actual_result);
        }
    }
}