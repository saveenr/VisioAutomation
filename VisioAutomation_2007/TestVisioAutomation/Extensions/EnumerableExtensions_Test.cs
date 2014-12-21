using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;

namespace TestVisioAutomation
{
    [TestClass]
    public class EnumerableExtensions : VisioAutomationTest
    {
        [TestMethod]
        public void Extensions_TestAsEnumerable()
        {
            this.Layers();
            this.Colors();
            this.Documents();
            this.Windows();
            this.Masters();
            this.Fonts();
            this.EnumeratePages();
            this.EnumerateShapes();
        }

        public void EnumerateShapes()
        {
            var page1 = GetNewPage();
            var app = page1.Application;

            // -------------------------------
            var a1 = page1.Shapes.AsEnumerable().ToList();
            Assert.AreEqual(0, a1.Count);

            var a2 = VA.Shapes.ShapeHelper.GetNestedShapes(page1.Shapes.AsEnumerable());
            Assert.AreEqual(0, a2.Count);

            // -------------------------------

            var s1 = page1.DrawRectangle(0, 0, 1, 1);
            var b1 = page1.Shapes.AsEnumerable().ToList();
            Assert.AreEqual(1, b1.Count);

            var b2 = VA.Shapes.ShapeHelper.GetNestedShapes(page1.Shapes.AsEnumerable());
            Assert.AreEqual(1, b2.Count);

            // -------------------------------

            var s2 = page1.DrawRectangle(1, 0, 2, 1);
            var s3 = page1.DrawRectangle(2, 0, 3, 1);
            var c1 = page1.Shapes.AsEnumerable().ToList();
            Assert.AreEqual(3, c1.Count);

            var c2 = VA.Shapes.ShapeHelper.GetNestedShapes(page1.Shapes.AsEnumerable());
            Assert.AreEqual(3, c2.Count);

            // -------------------------------

            var active_window = app.ActiveWindow;
            var selection = active_window.Selection;
            selection.DeselectAll();
            var g1 = VisioAutomationTest.SelectAndGroup(active_window, new[] { s2, s3 });

            var d1 = page1.Shapes.AsEnumerable().ToList();
            Assert.AreEqual(2, d1.Count);

            var d2 = VA.Shapes.ShapeHelper.GetNestedShapes(page1.Shapes.AsEnumerable());
            Assert.AreEqual(4, d2.Count);

            page1.Delete(0);
        }

        public void EnumeratePages()
        {
            var doc1 = this.GetNewDoc();
            var docpages = doc1.Pages;
            var page1 = docpages[1];
            var page2 = docpages.Add();
            var page3 = docpages.Add();

            page1.NameU = "P1";
            page2.NameU = "P2";
            page3.NameU = "P3";
            var pages = doc1.Pages;
            var expected = doc1.Pages.Cast<IVisio.Page>().ToList();
            var actual = doc1.Pages.AsEnumerable().ToList();

            Assert.AreEqual(expected.Count, actual.Count);
            Assert.AreEqual(pages[1].NameU, actual[0].NameU);
            Assert.AreEqual(pages[2].NameU, actual[1].NameU);
            Assert.AreEqual(pages[3].NameU, actual[2].NameU);

            doc1.Close(true);
        }

        public void Fonts()
        {
            var page1 = GetNewPage();
            var doc1 = page1.Document;
            var fonts = doc1.Fonts;

            var expects = fonts.Cast<IVisio.Font>().ToList();
            var actual = fonts.AsEnumerable().ToList();

            Assert.AreEqual(expects.Count, actual.Count);
            for (int i = 0; i < fonts.Count; i++)
            {
                Assert.AreEqual(fonts[i + 1].Name, actual[i].Name);
            }

            page1.Delete(0);
        }

        public void Masters()
        {
            var doc1 = this.GetNewDoc();
            var app = doc1.Application;
            var docs = app.Documents;

            var stencil = docs.OpenStencil("basic_u.vss");

            var masters = stencil.Masters;

            var actual = masters.AsEnumerable().ToList();
            for (int i = 0; i < masters.Count; i++)
            {
                Assert.AreEqual(masters[i + 1].NameU, actual[i].NameU);
            }

            doc1.Close(true);
        }

        public void Windows()
        {
            var doc1 = GetNewDoc();
            var app = doc1.Application;
            var windows = app.Windows;
            var actual = windows.AsEnumerable().ToList();
            for (int i = 0; i < windows.Count; i++)
            {
                var ex = windows[(short)(i + 1)];
                var ac = actual[i];
                Assert.AreEqual(ex.ID, ac.ID);
            }
            doc1.Close(true);
        }

        public void Documents()
        {
            var app = GetVisioApplication();
            var documents = app.Documents;
            var doc1 = documents.Add(string.Empty);
            var doc2 = documents.Add(string.Empty);
            var doc3 = documents.Add(string.Empty);

            doc1.Title = "D1";
            doc2.Title = "D2";
            doc3.Title = "D3";

            var actual = documents.AsEnumerable().ToList();
            for (int i = 0; i < documents.Count; i++)
            {
                Assert.AreEqual(documents[i + 1].Title, actual[i].Title);
            }

            doc1.Close(true);
            doc2.Close(true);
            doc3.Close(true);
        }

        public void Layers()
        {
            var doc1 = GetNewDoc();
            var page1 = doc1.Pages[1];

            var layers = page1.Layers;
            layers.Add("FOO");
            layers.Add("BAR");
            layers.Add("BEER");

            var actual = layers.AsEnumerable().ToList();
            for (int i = 0; i < layers.Count; i++)
            {
                var ex = layers[i+1];
                var ac = actual[i];
                Assert.AreEqual(ex.NameU, ac.NameU);
            }
            doc1.Close(true);
        }

        public void Colors()
        {
            var doc1 = GetNewDoc();
            var colors = doc1.Colors;
            var actual = colors.AsEnumerable().ToList();
            for (int i = 0; i < colors.Count; i++)
            {
                var expected_color = colors[i];
                var actual_color = actual[i];
                Assert.AreEqual(expected_color.Red, actual_color.Red);
                Assert.AreEqual(expected_color.Green, actual_color.Green);
                Assert.AreEqual(expected_color.Blue, actual_color.Blue);
            }
            doc1.Close(true);
        }

    }
}