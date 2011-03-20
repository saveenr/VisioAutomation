using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;
using VA = VisioAutomation;

namespace TestVisioAutomation
{
    [TestClass]
    public class PagesExtensions_Tests : VisioAutomationTest
    {
        [TestMethod]
        public void CreatePage()
        {
            var page1 = GetNewPage();
            var doc1 = page1.Document;

            int old_page_count = doc1.Pages.Count;

            page1.NameU = "A";

            var page2 = doc1.Pages.Add();
            Assert.AreEqual(old_page_count + 1, doc1.Pages.Count);
            page2.Name = "B";

            var page3 = doc1.Pages.Add();
            Assert.AreEqual(old_page_count + 2, doc1.Pages.Count);
            page3.Name = "C";

            short renum_pages = 1;
            page2.Delete(renum_pages);
            Assert.AreEqual(old_page_count + 1, doc1.Pages.Count);

            page3.Delete(renum_pages);
            Assert.AreEqual(old_page_count, doc1.Pages.Count);

            doc1.Close(true);
        }

        [TestMethod]
        public void ActivatePage()
        {
            var page1 = GetNewPage();
            var app = page1.Application;
            var doc1 = page1.Document;

            int old_page_count = doc1.Pages.Count;

            var page2 = doc1.Pages.Add();
            var page3 = doc1.Pages.Add();
            Assert.AreEqual(old_page_count + 2, doc1.Pages.Count);

            Assert.AreSame(page3, app.ActivePage);
            page2.Activate();
            Assert.AreSame(page2, app.ActivePage);
            page1.Activate();
            Assert.AreSame(page1, app.ActivePage);
            doc1.Close(true);
        }

        [TestMethod]
        public void ResizePage()
        {
            var page1 = GetNewPage();
            Assert.AreEqual(this.StandardPageSize, page1.GetSize());

            page1.SetSize(10, 5);
            Assert.AreEqual(new VA.Drawing.Size(10, 5), page1.GetSize());

            page1.SetSize(6, 3);
            Assert.AreEqual(new VA.Drawing.Size(6, 3), page1.GetSize());
            page1.Delete(0);
        }
    }
}