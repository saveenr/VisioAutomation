using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;
using System.Linq;

namespace TestVisioAutomation
{
    [TestClass]
    public class DocumentExtensions_Tests : VisioAutomationTest
    {
        [TestMethod]
        public void TestDocActivation()
        {
            var app = GetVisioApplication();
            var documents = app.Documents;
            var doc1 = documents.Add(string.Empty);
            var doc2 = documents.Add(string.Empty);
            var doc3 = documents.Add(string.Empty);

            doc1.Activate();
            Assert.AreEqual(doc1, app.ActiveDocument);
            doc2.Activate();
            Assert.AreEqual(doc2, app.ActiveDocument);
            doc3.Activate();
            Assert.AreEqual(doc3, app.ActiveDocument);
            doc1.Activate();
            Assert.AreEqual(doc1, app.ActiveDocument);

            doc1.Close(true);
            doc2.Close(true);
            doc3.Close(true);
        }

        [TestMethod]
        public void TestDocForceClosing()
        {
            var app = GetVisioApplication();
            var documents = app.Documents;
            int old_count = documents.Count;
            var doc1 = documents.Add(string.Empty);
            Assert.AreEqual(old_count + 1, documents.Count);
            var page1 = doc1.Pages[1];
            var s1 = page1.DrawRectangle(1, 1, 2, 2);
            doc1.Close(true);
            Assert.AreEqual(old_count, documents.Count);
        }
        
        [TestMethod]
        public void CheckDocumentColors()
        {
            var page1 = GetNewPage();
            var document = page1.Document;
            var colors1 = document.Colors;
            var colors = colors1.AsEnumerable().ToList();
            Assert.IsTrue(colors.Count >= 24);

            Assert.AreEqual("#000000", colors[0].ToColorRGB().ToWebColorString());
            Assert.AreEqual("#ffffff", colors[1].ToColorRGB().ToWebColorString());
            Assert.AreEqual("#ff0000", colors[2].ToColorRGB().ToWebColorString());
            Assert.AreEqual("#00ff00", colors[3].ToColorRGB().ToWebColorString());
            Assert.AreEqual("#0000ff", colors[4].ToColorRGB().ToWebColorString());
            Assert.AreEqual("#ffff00", colors[5].ToColorRGB().ToWebColorString());
            Assert.AreEqual("#ff00ff", colors[6].ToColorRGB().ToWebColorString());
            Assert.AreEqual("#00ffff", colors[7].ToColorRGB().ToWebColorString());
            Assert.AreEqual("#800000", colors[8].ToColorRGB().ToWebColorString());
            Assert.AreEqual("#008000", colors[9].ToColorRGB().ToWebColorString());
            Assert.AreEqual("#000080", colors[10].ToColorRGB().ToWebColorString());
            Assert.AreEqual("#808000", colors[11].ToColorRGB().ToWebColorString());
            Assert.AreEqual("#800080", colors[12].ToColorRGB().ToWebColorString());
            Assert.AreEqual("#008080", colors[13].ToColorRGB().ToWebColorString());
            Assert.AreEqual("#c0c0c0", colors[14].ToColorRGB().ToWebColorString());
            Assert.AreEqual("#e6e6e6", colors[15].ToColorRGB().ToWebColorString());
            Assert.AreEqual("#cdcdcd", colors[16].ToColorRGB().ToWebColorString());
            Assert.AreEqual("#b3b3b3", colors[17].ToColorRGB().ToWebColorString());
            Assert.AreEqual("#9a9a9a", colors[18].ToColorRGB().ToWebColorString());
            Assert.AreEqual("#808080", colors[19].ToColorRGB().ToWebColorString());
            Assert.AreEqual("#666666", colors[20].ToColorRGB().ToWebColorString());
            Assert.AreEqual("#4d4d4d", colors[21].ToColorRGB().ToWebColorString());
            Assert.AreEqual("#333333", colors[22].ToColorRGB().ToWebColorString());
            Assert.AreEqual("#1a1a1a", colors[23].ToColorRGB().ToWebColorString());

            var c1 = System.Drawing.Color.Red;
            var c2 = (System.Drawing.Color) colors[2].ToColorRGB();
            Assert.AreEqual(c1.R,c2.R);
            Assert.AreEqual(c1.G, c2.G);
            Assert.AreEqual(c1.B, c2.B);
            page1.Delete(0);
        }
    }
}