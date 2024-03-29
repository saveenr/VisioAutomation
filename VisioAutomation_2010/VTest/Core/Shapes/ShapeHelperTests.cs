using MUT=Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VTest.Core.Shapes
{
    [MUT.TestClass]
    public class ShapeHelperTests : Framework.VTest
    {
        [MUT.TestMethod]
        public void ShapeHelper_Test_GetNestedShapes_SingleShapeNoSubShapes()
        {
            // For a single shape with no subshapes, GetNestedShapes returns the single starting shape
            var page = this.GetNewPage();
            var shape0 = page.DrawRectangle(1, 1, 3, 3);

            var shapes = VisioAutomation.Shapes.ShapeHelper.GetNestedShapes(shape0);

            MUT.Assert.AreEqual(1,shapes.Count);
            MUT.Assert.IsTrue(shapes.Contains(shape0));
            page.Delete(0);
        }

        [MUT.TestMethod]
        public void ShapeHelper_Test_GetNestedShapes_GroupWithTwoSubShapes()
        {
            // group with two shapes
            var page = this.GetNewPage();
            var shape0 = page.DrawRectangle(0, 0, 1, 1);
            var shape1 = page.DrawRectangle(2, 0, 3, 1);

            var active_window = page.Application.ActiveWindow;
            var group = SelectAndGroup(active_window, new[] { shape0, shape1 });
            var shapes = VisioAutomation.Shapes.ShapeHelper.GetNestedShapes(group);

            MUT.Assert.AreEqual(3, shapes.Count);
            MUT.Assert.IsTrue(shapes.Contains(shape0));
            MUT.Assert.IsTrue(shapes.Contains(shape1));
            MUT.Assert.IsTrue(shapes.Contains(group));
            page.Delete(0);
        }

        [MUT.TestMethod]
        public void ShapeHelper_Test_GetNestedShapes_GroupWithSubGroups()
        {
            // group with subgroups
            var page = this.GetNewPage();
            var active_window = page.Application.ActiveWindow;
            
            var shape0 = page.DrawRectangle(0, 0, 1, 1);
            var shape1 = page.DrawRectangle(2, 0, 3, 1);

            var group0 = SelectAndGroup(active_window, new[] { shape0, shape1 });
            page.Application.ActiveWindow.DeselectAll();

            var shape2 = page.DrawRectangle(0, 3, 1, 4);
            var shape3 = page.DrawRectangle(2, 0, 5, 6);
            
            var group1 = SelectAndGroup(active_window, new[] { shape2, shape3 });
            page.Application.ActiveWindow.Selection.DeselectAll();
            
            var group2 = SelectAndGroup(active_window, new[] { group0, group1 });
            page.Application.ActiveWindow.Selection.DeselectAll();

            var shapes = VisioAutomation.Shapes.ShapeHelper.GetNestedShapes(group2);

            MUT.Assert.AreEqual(7, shapes.Count);
            MUT.Assert.IsTrue(shapes.Contains(shape0));
            MUT.Assert.IsTrue(shapes.Contains(shape1));
            MUT.Assert.IsTrue(shapes.Contains(shape2));
            MUT.Assert.IsTrue(shapes.Contains(shape3));
            MUT.Assert.IsTrue(shapes.Contains(group0));
            MUT.Assert.IsTrue(shapes.Contains(group1));
            page.Delete(0);
        }
    }
}