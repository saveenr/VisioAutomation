using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;
using TREE = VisioAutomation.Models.Layouts.Tree;
using VA = VisioAutomation;

namespace VisioAutomation_Tests.Models.TreeLayout.Layouts
{
    [TestClass]
    public class Tree_Tests : VisioAutomationTest
    {
        [TestMethod]
        public void TreeLayout_SingleNode()
        {
            // Verify that a tree with a single node can be drawn
            var t = new TREE.Drawing();
            t.Root = new TREE.Node("Root");

            var app = this.GetVisioApplication();
            var doc = this.GetNewDoc();
            var page = app.ActivePage;

            t.Render(page);

            AssertUtil.AreEqual((3.0, 1.5), VisioAutomationTest.GetPageSize(page), 0.05);

            doc.Close(true);
        }

        [TestMethod]
        public void TreeLayout_MultiNode()
        {
            // Verify that a tree with multiple nodes can be drawn
            // Note that  the DefaultNodeSize option is being used

            var t = new TREE.Drawing();

            t.Root = new TREE.Node("Root");

            var na = new TREE.Node("A");
            var nb = new TREE.Node("B");

            var na1 = new TREE.Node("A1");
            var na2 = new TREE.Node("A2");

            var nb1 = new TREE.Node("B1");
            var nb2 = new TREE.Node("B2");

            t.Root.Children.Add(na);
            t.Root.Children.Add(nb);

            na.Children.Add(na1);
            na.Children.Add(na2);

            nb.Children.Add(nb1);
            nb1.Children.Add(nb2);

            t.LayoutOptions.DefaultNodeSize = new VA.Geometry.Size(1, 1);

            var app = this.GetVisioApplication();
            var doc = this.GetNewDoc();
            var page = app.ActivePage;

            t.Render(page);

            AssertUtil.AreEqual((5.25, 8.0), VisioAutomationTest.GetPageSize(page), 0.05);

            Assert.AreEqual("Root", t.Root.VisioShape.Text);
            Assert.AreEqual("A", na.VisioShape.Text);
            Assert.AreEqual("B", nb.VisioShape.Text);

            Assert.AreEqual("A1", na1.VisioShape.Text);
            Assert.AreEqual("A2", na2.VisioShape.Text);

            Assert.AreEqual("B1", nb1.VisioShape.Text);
            Assert.AreEqual("B2", nb2.VisioShape.Text);

            doc.Close(true);
        }
    }
}