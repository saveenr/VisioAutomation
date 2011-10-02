using System;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using System.Linq;
using BL = VisioAutomation.Layout.BoxLayout;

namespace TestVisioAutomation
{
    [TestClass]
    public class BoxLayoutTests : VisioAutomationTest
    {
        [TestMethod]
        public void Test_empty()
        {
            var layout = new VA.Layout.BoxLayout.BoxLayout<object>();
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

            int x = 1;
        }

        [TestMethod]
        public void Test_single_node()
        {
            var layout = new VA.Layout.BoxLayout.BoxLayout<object>();
            var root = layout.Root;
            var n1 = root.AddNode(10, 5);
            layout.PerformLayout();
            double delta = 0.00000001;
            AssertX.AreEqual(0,0,10,5, n1.ReservedRectangle,delta);
            AssertX.AreEqual(0, 0, 10, 5, n1.Rectangle, delta);

            AssertX.AreEqual(0, 0, 10, 5, root.ReservedRectangle, delta);
            AssertX.AreEqual(0, 0, 10, 5, root.Rectangle, delta);
            
        }

        [TestMethod]
        public void Test_single_node_padding()
        {
            var layout = new VA.Layout.BoxLayout.BoxLayout<object>();
            var root = layout.Root;
            var n1 = root.AddNode(10, 5);

            root.Padding = 1.0;
            layout.PerformLayout();
            double delta = 0.00000001;
            AssertX.AreEqual(1.0, 1.0, 11, 6, n1.ReservedRectangle, delta);
            AssertX.AreEqual(1.0, 1.0, 11, 6, n1.Rectangle, delta);


        }

        [TestMethod]
        public void Test_two_nodes()
        {
            var layout = new VA.Layout.BoxLayout.BoxLayout<object>();
            var root = layout.Root;
            var n1 = root.AddNode(1, 2);
            var n2 = root.AddNode(2, 3);

            root.Padding = 1.0;
            layout.PerformLayout();
            double delta = 0.00000001;
            AssertX.AreEqual(1, 1, 2, 3, n1.ReservedRectangle, delta);
            AssertX.AreEqual(1, 4, 3, 7, n1.Rectangle, delta);
        }


    }
}