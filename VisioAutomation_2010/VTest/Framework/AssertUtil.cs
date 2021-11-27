using MUT=Microsoft.VisualStudio.TestTools.UnitTesting;
using VADRAW = VisioAutomation.Core;

namespace VTest
{
    public static class AssertUtil
    {
        public static void FileExists(string filename)
        {
            MUT.Assert.IsTrue(System.IO.File.Exists(filename));
        }

        public static void AreEqual( (double x, double y) expected, VADRAW.Point actual, double delta)
        {
            MUT.Assert.AreEqual(expected.x, actual.X, delta);
            MUT.Assert.AreEqual(expected.y, actual.Y, delta);
        }

        public static void AreEqual( (double left, double bottom, double right, double top) expected, VADRAW.Rectangle actual_rect, double delta)
        {
            MUT.Assert.AreEqual(expected.left, actual_rect.Left, delta);
            MUT.Assert.AreEqual(expected.bottom, actual_rect.Bottom, delta);
            MUT.Assert.AreEqual(expected.right, actual_rect.Right, delta);
            MUT.Assert.AreEqual(expected.top, actual_rect.Top, delta);
        }

        public static void AreEqual( (double width, double height) expected, VADRAW.Size actual_size, double delta)
        {
            MUT.Assert.AreEqual(expected.width, actual_size.Width, delta);
            MUT.Assert.AreEqual(expected.height, actual_size.Height, delta);
        }

        public static void AreEqual<TResult>( (string formula, TResult result) expected, (string formula, TResult result) actual)
        {
            MUT.Assert.AreEqual(expected.formula, actual.formula);
            MUT.Assert.AreEqual(expected.result, actual.result);
        }

        public static void OneOf<T>(T[] expected, T actual)
        {
            bool match = false;
            foreach (var value in expected)
            {
                match = true;
                break;
            }
            if (!match)
            {
                string a = string.Join(",", expected);
                string b = string.Format("Expected one of ({0}), Actual is ({1}).", a, actual);
                MUT.Assert.Fail(b);
            }
        }
    }
}