using Microsoft.VisualStudio.TestTools.UnitTesting;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace TestVisioAutomation
{
    [TestClass]
    public class ShapeHelper_Tests : VisioAutomationTest
    {
        [TestMethod]
        public void ShapeHelper_Test_GetNestedShapes_SingleShapeNoSubShapes()
        {
            // For a single shape with no subshapes, GetNestedShapes returns the single starting shape
            var page = this.GetNewPage();
            var shape0 = page.DrawRectangle(1, 1, 3, 3);

            var shapes = VA.Shapes.ShapeHelper.GetNestedShapes(shape0);

            Assert.AreEqual(1,shapes.Count);
            Assert.IsTrue(shapes.Contains(shape0));
            page.Delete(0);
        }

        [TestMethod]
        public void ShapeHelper_Test_GetNestedShapes_GroupWithTwoSubShapes()
        {
            // group with two shapes
            var page = this.GetNewPage();
            var shape0 = page.DrawRectangle(0, 0, 1, 1);
            var shape1 = page.DrawRectangle(2, 0, 3, 1);

            var active_window = page.Application.ActiveWindow;
            var group = VisioAutomationTest.SelectAndGroup(active_window, new[] { shape0, shape1 });
            var shapes = VA.Shapes.ShapeHelper.GetNestedShapes(group);

            Assert.AreEqual(3, shapes.Count);
            Assert.IsTrue(shapes.Contains(shape0));
            Assert.IsTrue(shapes.Contains(shape1));
            Assert.IsTrue(shapes.Contains(group));
            page.Delete(0);
        }

        [TestMethod]
        public void ShapeHelper_Test_GetNestedShapes_GroupWithSubGroups()
        {
            // group with subgroups
            var page = this.GetNewPage();
            var active_window = page.Application.ActiveWindow;
            
            var shape0 = page.DrawRectangle(0, 0, 1, 1);
            var shape1 = page.DrawRectangle(2, 0, 3, 1);

            var group0 = VisioAutomationTest.SelectAndGroup(active_window, new[] { shape0, shape1 });
            page.Application.ActiveWindow.DeselectAll();

            var shape2 = page.DrawRectangle(0, 3, 1, 4);
            var shape3 = page.DrawRectangle(2, 0, 5, 6);
            
            var group1 = VisioAutomationTest.SelectAndGroup(active_window, new[] { shape2, shape3 });
            page.Application.ActiveWindow.Selection.DeselectAll();
            
            var group2 = VisioAutomationTest.SelectAndGroup(active_window, new[] { group0, group1 });
            page.Application.ActiveWindow.Selection.DeselectAll();

            var shapes = VA.Shapes.ShapeHelper.GetNestedShapes(group2);

            Assert.AreEqual(7, shapes.Count);
            Assert.IsTrue(shapes.Contains(shape0));
            Assert.IsTrue(shapes.Contains(shape1));
            Assert.IsTrue(shapes.Contains(shape2));
            Assert.IsTrue(shapes.Contains(shape3));
            Assert.IsTrue(shapes.Contains(group0));
            Assert.IsTrue(shapes.Contains(group1));
            page.Delete(0);
        }
    }
}