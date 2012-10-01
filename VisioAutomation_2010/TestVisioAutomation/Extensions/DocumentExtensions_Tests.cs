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
    }
}