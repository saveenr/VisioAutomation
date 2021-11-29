using VTest.Framework;
using MUT=Microsoft.VisualStudio.TestTools.UnitTesting;
using VABOX = VisioAutomation.Models.Layouts.Box;

namespace VTest.Models
{
    [MUT.TestClass]
    public class Layout_BoxTests : Framework.VTest
    {
        [MUT.TestMethod]
        public void BoxLayout_Test_empty()
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
        public void BoxLayout_Test_single_node()
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
        public void BoxLayout_Test_single_node_padding()
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
    }
}