using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace TestVisioAutomation
{
    [TestClass]
    public class TreeTests : VisioAutomationTest
    {

        [TestMethod]
        public void DrawTree1Node()
        {
            var t = new VA.Layout.Tree.Drawing();
            t.Root = new VA.Layout.Tree.Node("Root");

            var app = this.GetVisioApplication();
            var doc = this.GetNewDoc();
            var page = app.ActivePage;

            t.Render(page);

            TestUtil.AreEqual(3.0, 1.5, page.GetSize(), 0.05);

            doc.Close(true);

        }


        [TestMethod]
        public void DrawTreeMultiNode()
        {
            var t = new VA.Layout.Tree.Drawing();
            t.Root = new VA.Layout.Tree.Node("Root");

            var na = new VA.Layout.Tree.Node("A");
            var nb = new VA.Layout.Tree.Node("B");

            var na1 = new VA.Layout.Tree.Node("A1");
            var na2 = new VA.Layout.Tree.Node("A2");

            var nb1 = new VA.Layout.Tree.Node("B1");
            var nb2 = new VA.Layout.Tree.Node("B2");

            t.Root.Children.Add(na);
            t.Root.Children.Add(nb);

            na.Children.Add(na1);
            na.Children.Add(na2);

            nb.Children.Add(nb1);
            nb1.Children.Add(nb2);

            var app = this.GetVisioApplication();
            var doc = this.GetNewDoc();
            var page = app.ActivePage;

            t.Render(page);

            TestUtil.AreEqual(8.25, 6.0, page.GetSize(), 0.05);

            Assert.AreEqual("Root", t.Root.VisioShape.Text);
            Assert.AreEqual("A", na.VisioShape.Text);
            Assert.AreEqual("B", nb.VisioShape.Text);

            Assert.AreEqual("A1", na1.VisioShape.Text);
            Assert.AreEqual("A2", na2.VisioShape.Text);

            Assert.AreEqual("B1", nb1.VisioShape.Text);
            Assert.AreEqual("B2", nb2.VisioShape.Text);

            doc.Close(true);

        }


        [TestMethod]
        public void DrawTreeMultiNode2()
        {
            var t = new VA.Layout.Tree.Drawing();

            t.Root = new VA.Layout.Tree.Node("Root");

            var na = new VA.Layout.Tree.Node("A");
            var nb = new VA.Layout.Tree.Node("B");

            var na1 = new VA.Layout.Tree.Node("A1");
            var na2 = new VA.Layout.Tree.Node("A2");

            var nb1 = new VA.Layout.Tree.Node("B1");
            var nb2 = new VA.Layout.Tree.Node("B2");

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

            TestUtil.AreEqual(5.25, 8.0, page.GetSize(), 0.05);

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