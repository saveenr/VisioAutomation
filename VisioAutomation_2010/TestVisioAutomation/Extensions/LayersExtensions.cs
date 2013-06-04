using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;

namespace TestVisioAutomation
{
    [TestClass]
    public class EnumerableExtensions : VisioAutomationTest
    {
        [TestMethod]
        public void TestAsEnumerable()
        {
            this.Layers();
            this.Colors();
            this.Documents();
            this.Windows();
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