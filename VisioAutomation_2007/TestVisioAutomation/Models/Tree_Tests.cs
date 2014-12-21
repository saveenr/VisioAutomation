using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using TREEMODEL = VisioAutomation.Models.Tree;

namespace TestVisioAutomation
{
    [TestClass]
    public class Tree_Tests : VisioAutomationTest
    {
        [TestMethod]
        public void TreeLayout_SingleNode()
        {
            // Verify that a tree with a single node can be drawn
            var t = new TREEMODEL.Drawing();
            t.Root = new TREEMODEL.Node("Root");

            var app = this.GetVisioApplication();
            var doc = this.GetNewDoc();
            var page = app.ActivePage;

            t.Render(page);

            AssertVA.AreEqual(3.0, 1.5, VisioAutomationTest.GetPageSize(page), 0.05);

            doc.Close(true);
        }

        [TestMethod]
        public void TreeLayout_MultiNode()
        {
            // Verify that a tree with multiple nodes can be drawn
            // Note that  the DefaultNodeSize option is being used

            var t = new TREEMODEL.Drawing();

            t.Root = new TREEMODEL.Node("Root");

            var na = new TREEMODEL.Node("A");
            var nb = new TREEMODEL.Node("B");

            var na1 = new TREEMODEL.Node("A1");
            var na2 = new TREEMODEL.Node("A2");

            var nb1 = new TREEMODEL.Node("B1");
            var nb2 = new TREEMODEL.Node("B2");

            t.Root.Children.Add(na);
            t.Root.Children.Add(nb);

            na.Children.Add(na1);
            na.Children.Add(na2);

            nb.Children.Add(nb1);
            nb1.Children.Add(nb2);

            t.LayoutOptions.DefaultNodeSize = new VA.Drawing.Size(1, 1);

            var app = this.GetVisioApplication();
            var doc = this.GetNewDoc();
            var page = app.ActivePage;

            t.Render(page);

            AssertVA.AreEqual(5.25, 8.0, VisioAutomationTest.GetPageSize(page), 0.05);

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