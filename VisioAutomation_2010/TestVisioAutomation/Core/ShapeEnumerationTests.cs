using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;
using VA = VisioAutomation;

namespace TestVisioAutomation
{
    [TestClass]
    public class ShapeEnumerationTests : VisioAutomationTest
    {
        [TestMethod]
        public void EnumerateShapes()
        {
            var page1 = GetNewPage();
            var app = page1.Application;

            // -------------------------------
            var a1 = page1.Shapes.AsEnumerable().ToList();
            Assert.AreEqual(0, a1.Count);

            var a2 = VA.ShapeHelper.GetNestedShapes(page1.Shapes.AsEnumerable());
            Assert.AreEqual(0, a2.Count);

            // -------------------------------

            var s1 = page1.DrawRectangle(0, 0, 1, 1);
            var b1 = page1.Shapes.AsEnumerable().ToList();
            Assert.AreEqual(1, b1.Count);

            var b2 = VA.ShapeHelper.GetNestedShapes(page1.Shapes.AsEnumerable());
            Assert.AreEqual(1, b2.Count);

            // -------------------------------

            var s2 = page1.DrawRectangle(1, 0, 2, 1);
            var s3 = page1.DrawRectangle(2, 0, 3, 1);
            var c1 = page1.Shapes.AsEnumerable().ToList();
            Assert.AreEqual(3, c1.Count);

            var c2 = VA.ShapeHelper.GetNestedShapes(page1.Shapes.AsEnumerable());
            Assert.AreEqual(3, c2.Count);

            // -------------------------------

            var active_window = app.ActiveWindow;
            var selection = active_window.Selection;
            selection.DeselectAll();
            var g1 = VisioAutomationTest.SelectAndGroup(active_window, new[] { s2, s3 });

            var d1 = page1.Shapes.AsEnumerable().ToList();
            Assert.AreEqual(2, d1.Count);

            var d2 = VA.ShapeHelper.GetNestedShapes(page1.Shapes.AsEnumerable());
            Assert.AreEqual(4, d2.Count);

            page1.Delete(0);
        }
    }
}