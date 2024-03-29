using MUT=Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;

namespace VTest.Core.Extensions
{
    [MUT.TestClass]
    public class PageTests : Framework.VTest
    {
        [MUT.TestMethod]
        public void Page_CreatePage()
        {
            var page1 = this.GetNewPage();
            var doc1 = page1.Document;

            int old_page_count = doc1.Pages.Count;

            page1.NameU = "A";

            var page2 = doc1.Pages.Add();
            MUT.Assert.AreEqual(old_page_count + 1, doc1.Pages.Count);
            page2.Name = "B";

            var page3 = doc1.Pages.Add();
            MUT.Assert.AreEqual(old_page_count + 2, doc1.Pages.Count);
            page3.Name = "C";

            short renum_pages = 1;
            page2.Delete(renum_pages);
            MUT.Assert.AreEqual(old_page_count + 1, doc1.Pages.Count);

            page3.Delete(renum_pages);
            MUT.Assert.AreEqual(old_page_count, doc1.Pages.Count);

            doc1.Close(true);
        }
    }
}