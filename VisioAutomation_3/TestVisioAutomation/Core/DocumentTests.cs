using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace TestVisioAutomation
{
    [TestClass]
    public class DocumentTests : VisioAutomationTest
    {
        [TestMethod]
        public void SwitchPages()
        {
            var app = this.GetVisioApplication();

            var documents = app.Documents;
            int old_doc_count = documents.Count;

            var doc1 = this.GetNewDoc();
            Assert.AreEqual(documents.Count, old_doc_count + 1);
            Assert.AreEqual(doc1.Pages.Count, 1);
            var page1 = doc1.Pages[1];
            Assert.AreEqual(app.ActivePage, page1);

            var page2 = doc1.Pages.Add();
            page2.Background = 0;
            page2.SetSize(this.StandardPageSize);

            Assert.AreEqual(app.ActivePage, page2);
            page1.Activate();
            Assert.AreEqual(app.ActivePage, page1);
            page2.Activate();
            Assert.AreEqual(app.ActivePage, page2);
            doc1.Close(true);
        }

        [TestMethod]
        public void SetSize()
        {
            var app = this.GetVisioApplication();
            var documents = app.Documents;
            var doc1 = this.GetNewDoc();
            var page1 = doc1.Pages[1];
            page1.SetSize(this.StandardPageSize);
            var page_size = page1.GetSize();
            Assert.AreEqual(page_size, this.StandardPageSize);
            doc1.Close(true);
        }
    }
}