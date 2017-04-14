using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Scripting.Models;
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

        public static void AreEqual( (double x, double y) expected, VADRAW.Point actual, double delta)
        {
            Assert.AreEqual(expected.x, actual.X, delta);
            Assert.AreEqual(expected.y, actual.Y, delta);
        }

        public static void AreEqual( (double left, double bottom, double right, double top) expected, VADRAW.Rectangle actual_rect, double delta)
        {
            Assert.AreEqual(expected.left, actual_rect.Left, delta);
            Assert.AreEqual(expected.bottom, actual_rect.Bottom, delta);
            Assert.AreEqual(expected.right, actual_rect.Right, delta);
            Assert.AreEqual(expected.top, actual_rect.Top, delta);
        }

        public static void AreEqual( (double width, double height) expected, VADRAW.Size actual_size, double delta)
        {
            Assert.AreEqual(expected.width, actual_size.Width, delta);
            Assert.AreEqual(expected.height, actual_size.Height, delta);
        }

        public static void AssertSnap((double x, double y) expected, SnappingGrid snapgrid, (double x, double y) input, double delta)
        {
            var snapped = snapgrid.Snap(input.x, input.y);
            Assert.AreEqual(expected.x, snapped.X, delta);
            Assert.AreEqual(expected.y, snapped.Y, delta);
        }

        public static void AreEqual<TResult>( (string formula, TResult result) expected, (string formula, TResult result) actual)
        {
            Assert.AreEqual(expected.formula, actual.formula);
            Assert.AreEqual(expected.result, actual.result);
        }
    }
}