using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;
using VA = VisioAutomation;
using IVisio= Microsoft.Office.Interop.Visio;

namespace TestVisioAutomation
{
    [TestClass]
    public class Pages_Tests : VisioAutomationTest
    {
        [TestMethod]
        public void Page_CreatePage()
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
    }
}