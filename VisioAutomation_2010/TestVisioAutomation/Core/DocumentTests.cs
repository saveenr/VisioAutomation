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
            VisioAutomationTest.SetPageSize(page2, this.StandardPageSize);

            var active_window = app.ActiveWindow;
            Assert.AreEqual(app.ActivePage, page2);
            active_window.Page = page1;
            Assert.AreEqual(app.ActivePage, page1);
            active_window.Page = page2;
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
            VisioAutomationTest.SetPageSize(page1, this.StandardPageSize);
            var page_size = VisioAutomationTest.GetPageSize(page1);
            Assert.AreEqual(page_size, this.StandardPageSize);
            doc1.Close(true);
        }
    }
}