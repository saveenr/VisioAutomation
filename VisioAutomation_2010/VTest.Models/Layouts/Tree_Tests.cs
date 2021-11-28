using VisioAutomation.Extensions;
using VTest.Framework;
using MUT=Microsoft.VisualStudio.TestTools.UnitTesting;
using VATREE = VisioAutomation.Models.Layouts.Tree;
using VA = VisioAutomation;

namespace VTest.Models.Layouts
{
    [MUT.TestClass]
    public class Tree_Tests : Framework.VTest
    {
        [MUT.TestMethod]
        public void TreeLayout_SingleNode()
        {
            // Verify that a tree with a single node can be drawn
            var t = new VATREE.Drawing();
            t.Root = new VATREE.Node("Root");

            var app = this.GetVisioApplication();
            var doc = this.GetNewDoc();
            var page = app.ActivePage;

            t.Render(page);

            AssertUtil.AreEqual((3.0, 1.5), Framework.VTest.GetPageSize(page), 0.05);

            doc.Close(true);
        }

        [MUT.TestMethod]
        public void TreeLayout_MultiNode()
        {
            // Verify that a tree with multiple nodes can be drawn
            // Note that  the DefaultNodeSize option is being used

            var t = new VATREE.Drawing();

            t.Root = new VATREE.Node("Root");

            var na = new VATREE.Node("A");
            var nb = new VATREE.Node("B");

            var na1 = new VATREE.Node("A1");
            var na2 = new VATREE.Node("A2");

            var nb1 = new VATREE.Node("B1");
            var nb2 = new VATREE.Node("B2");

            t.Root.Children.Add(na);
            t.Root.Children.Add(nb);

            na.Children.Add(na1);
            na.Children.Add(na2);

            nb.Children.Add(nb1);
            nb1.Children.Add(nb2);

            t.LayoutOptions.DefaultNodeSize = new VA.Core.Size(1, 1);

            var app = this.GetVisioApplication();
            var doc = this.GetNewDoc();
            var page = app.ActivePage;

            t.Render(page);

            AssertUtil.AreEqual((5.25, 8.0), Framework.VTest.GetPageSize(page), 0.05);

            MUT.Assert.AreEqual("Root", t.Root.VisioShape.Text);
            MUT.Assert.AreEqual("A", na.VisioShape.Text);
            MUT.Assert.AreEqual("B", nb.VisioShape.Text);

            MUT.Assert.AreEqual("A1", na1.VisioShape.Text);
            MUT.Assert.AreEqual("A2", na2.VisioShape.Text);

            MUT.Assert.AreEqual("B1", nb1.VisioShape.Text);
            MUT.Assert.AreEqual("B2", nb2.VisioShape.Text);

            doc.Close(true);
        }
    }
}