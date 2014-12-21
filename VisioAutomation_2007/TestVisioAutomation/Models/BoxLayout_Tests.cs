using Microsoft.VisualStudio.TestTools.UnitTesting;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using BoxL = VisioAutomation.Models.BoxLayout;

namespace TestVisioAutomation
{
    [TestClass]
    public class BoxLayout_Tests : VisioAutomationTest
    {
        [TestMethod]
        public void BoxLayout_Test_empty()
        {
            var layout = new BoxL.BoxLayout();
            layout.Root = new BoxL.Container(BoxL.Direction.BottomToTop);
            Assert.IsNotNull(layout.Root);

            bool thrown = false;
            try
            {
                layout.PerformLayout();

            }
            catch (VA.AutomationException)
            {
                thrown = true;
            }

            if (!thrown)
            {
                Assert.Fail();
            }
        }

        [TestMethod]
        public void BoxLayout_Test_single_node()
        {
            var layout = new BoxL.BoxLayout();
            layout.Root = new BoxL.Container(BoxL.Direction.BottomToTop);
            var root = layout.Root;
            root.PaddingBottom = 0.0;
            root.PaddingLeft= 0.0;
            root.PaddingRight= 0.0;
            root.PaddingTop= 0.0;
            var n1 = root.AddBox(10, 5);
            layout.PerformLayout();
            double delta = 0.00000001;

            AssertVA.AreEqual(0, 0, 10, 5, n1.Rectangle, delta);
            AssertVA.AreEqual(0, 0, 10, 5, root.Rectangle, delta);          
        }

        [TestMethod]
        public void BoxLayout_Test_single_node_padding()
        {
            var layout = new BoxL.BoxLayout();
            layout.Root = new BoxL.Container(BoxL.Direction.BottomToTop);
            var root = layout.Root;
            var n1 = root.AddBox(10, 5);

            root.PaddingBottom = 1.0;
            root.PaddingLeft = 1.0;
            root.PaddingRight = 1.0;
            root.PaddingTop = 1.0;

            layout.PerformLayout();
            double delta = 0.00000001;
            AssertVA.AreEqual(1.0, 1.0, 11, 6, n1.Rectangle, delta);
        }
    }
}