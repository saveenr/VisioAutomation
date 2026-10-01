using VTest.Framework;
using MUT=Microsoft.VisualStudio.TestTools.UnitTesting;
using VABOX = VisioAutomation.Models.Layouts.Box;

namespace VTest.Models
{
    [MUT.TestClass]
    public class BoxLayoutTests : Framework.VTest
    {
        [MUT.TestMethod]
        public void EmptyContainer_PerformLayoutThrowsArgumentException()
        {
            var layout = new VABOX.BoxLayout();
            layout.Root = new VABOX.Container(VABOX.Direction.BottomToTop);
            MUT.Assert.IsNotNull(layout.Root);

            bool thrown = false;
            try
            {
                layout.PerformLayout();

            }
            catch (System.ArgumentException)
            {
                thrown = true;
            }

            if (!thrown)
            {
                MUT.Assert.Fail();
            }
        }

        [MUT.TestMethod]
        public void SingleBoxNoPadding_RectangleMatchesBoxBounds()
        {
            var layout = new VABOX.BoxLayout();
            layout.Root = new VABOX.Container(VABOX.Direction.BottomToTop);
            var root = layout.Root;
            root.PaddingBottom = 0.0;
            root.PaddingLeft= 0.0;
            root.PaddingRight= 0.0;
            root.PaddingTop= 0.0;
            var n1 = root.AddBox(10, 5);
            layout.PerformLayout();
            double delta = 0.00000001;

            AssertUtil.AreEqual((0, 0, 10, 5), n1.Rectangle, delta);
            AssertUtil.AreEqual((0, 0, 10, 5), root.Rectangle, delta);          
        }

        [MUT.TestMethod]
        public void SingleBoxWithUniformPadding_BoxIsOffsetByPadding()
        {
            var layout = new VABOX.BoxLayout();
            layout.Root = new VABOX.Container(VABOX.Direction.BottomToTop);
            var root = layout.Root;
            var n1 = root.AddBox(10, 5);

            root.PaddingBottom = 1.0;
            root.PaddingLeft = 1.0;
            root.PaddingRight = 1.0;
            root.PaddingTop = 1.0;

            layout.PerformLayout();
            double delta = 0.00000001;
            AssertUtil.AreEqual((1.0, 1.0, 11, 6), n1.Rectangle, delta);
        }

        [MUT.TestMethod]
        public void NestedRightToLeftContainer_ChildrenPlacedRelativeToContainerX()
        {
            // The right-to-left container is placed at origin (0, 3), where origin.Y != origin.X,
            // so using the Y coordinate for the starting X would put the children at the wrong place.
            var layout = new VABOX.BoxLayout();
            layout.Root = new VABOX.Container(VABOX.Direction.BottomToTop);
            var root = layout.Root;
            root.PaddingBottom = 0.0;
            root.PaddingLeft = 0.0;
            root.PaddingRight = 0.0;
            root.PaddingTop = 0.0;
            root.ChildSpacing = 0.0;

            root.AddBox(4, 3);

            var rtl = root.AddContainer(VABOX.Direction.RightToLeft);
            rtl.PaddingBottom = 0.0;
            rtl.PaddingLeft = 0.0;
            rtl.PaddingRight = 0.0;
            rtl.PaddingTop = 0.0;
            rtl.ChildSpacing = 0.0;
            var first = rtl.AddBox(2, 1);
            var second = rtl.AddBox(1, 1);

            layout.PerformLayout();
            double delta = 0.00000001;

            // Right to left: the first child sits at the right edge of the container (x = 3)
            AssertUtil.AreEqual((1, 3, 3, 4), first.Rectangle, delta);
            AssertUtil.AreEqual((0, 3, 1, 4), second.Rectangle, delta);
        }
    }
}