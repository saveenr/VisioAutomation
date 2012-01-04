using System;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Drawing;
using VisioAutomation.Extensions;
using VisioAutomation.Layout.BoxLayout;
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
            var layout = new VA.Layout.BoxLayout.BoxLayout();
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
        public void Test_single_node()
        {
            var layout = new VA.Layout.BoxLayout.BoxLayout();
            var root = layout.Root;
            var n1 = root.AddBox(10, 5);
            layout.PerformLayout();
            double delta = 0.00000001;
            AssertVA.AreEqual(0, 0, 10, 5, n1.Rectangle, delta);

            AssertVA.AreEqual(0, 0, 10, 5, root.Rectangle, delta);
            
        }

        [TestMethod]
        public void Test_single_node_padding()
        {
            var layout = new VA.Layout.BoxLayout.BoxLayout();
            var root = layout.Root;
            var n1 = root.AddBox(10, 5);

            root.Padding = 1.0;
            layout.PerformLayout();
            double delta = 0.00000001;
            AssertVA.AreEqual(1.0, 1.0, 11, 6, n1.Rectangle, delta);


        }

        [TestMethod]
        public void Test_two_nodes_1()
        {
            var layout = new VA.Layout.BoxLayout.BoxLayout();
            var root = layout.Root;
            root.Data = "Root";
            var n1 = root.AddBox(1, 2);
            n1.Data = "n1";
            var n2 = root.AddBox(2, 3);
            n2.Data = "n2";

            root.Padding = 1.0;
            layout.PerformLayout();
            double delta = 0.00000001;

            AssertVA.AreEqual(1, 1, 2, 3, n1.Rectangle, delta);
            AssertVA.AreEqual(1, 3, 3, 6, n2.Rectangle, delta);
        }


        [TestMethod]
        public void Test_two_nodes_2()
        {
            var layout = new VA.Layout.BoxLayout.BoxLayout();
            var root = layout.Root;
            root.Data = "Root";
            var n1 = root.AddBox(1, 2);
            n1.Data = "n1";
            n1.AlignmentHorizontal = AlignmentHorizontal.Right;
            var n2 = root.AddBox(2, 3);
            n2.Data = "n2";

            root.Padding = 1.0;
            layout.PerformLayout();
            double delta = 0.00000001;


            var doc = draw_layout(layout);
            //doc.Close(true);

            AssertVA.AreEqual(2, 1, 3, 3, n1.Rectangle, delta);
            AssertVA.AreEqual(1, 3, 3, 6, n2.Rectangle, delta);

        }

        private IVisio.Document draw_layout(BoxLayout layout)
        {
            var app = this.GetVisioApplication();
            var doc = this.GetNewDoc();
            var page = app.ActivePage;
            var dom = new VA.DOM.Document();

            var list = new List<VA.Layout.BoxLayout.Node>();
            list.Add(layout.Root);
            list.AddRange(layout.Nodes);
            foreach (var node in list)
            {
                var dom_n1 = dom.DrawRectangle(node.Rectangle);
                dom_n1.ShapeCells.FillForegnd = "rgb(255,190,0)";

                dom_n1.Text = string.Format("{0}\n{1}", node.Data, node.Rectangle.ToString());
            }

            dom.Render(page);
            page.ResizeToFitContents(1, 1);
            return doc;
        }


        [TestMethod]
        public void Test_two_nodes_3()
        {
            var layout = new VA.Layout.BoxLayout.BoxLayout();
            var root = layout.Root;
            root.Data = "Root";
            root.Direction = LayoutDirection.Horizonal;
            var n1 = root.AddBox(1, 2);
            n1.Data = "n1";

            n1.AlignmentVertical = AlignmentVertical.Top;
            var n2 = root.AddBox(2, 3);
            n2.Data = "n2";

            root.Padding = 1.0;
            layout.PerformLayout();
            double delta = 0.00000001;

            AssertVA.AreEqual(1, 2, 2, 4, n1.Rectangle, delta);

            AssertVA.AreEqual(2, 1, 4, 4, n2.Rectangle, delta);
        }

        [TestMethod]
        public void Test_two_nodes_4()
        {
            var layout = new VA.Layout.BoxLayout.BoxLayout();
            layout.LayoutOptions.DirectionHorizontal = DirectionHorizontal.RightToLeft;
            var root = layout.Root;

            root.Data = "Root";
            root.Direction = LayoutDirection.Horizonal;
            var n1 = root.AddBox(1, 2);
            n1.Data = "n1";

            n1.AlignmentVertical = AlignmentVertical.Top;
            var n2 = root.AddBox(2, 3);
            n2.Data = "n2";

            root.Padding = 1.0;
            layout.PerformLayout();
            double delta = 0.00000001;

            var doc = draw_layout(layout);

            //AssertX.AreEqual(1, 2, 2, 4, n1.Rectangle, delta);

            //AssertX.AreEqual(2, 1, 4, 4, n2.Rectangle, delta);

        }

    }
}