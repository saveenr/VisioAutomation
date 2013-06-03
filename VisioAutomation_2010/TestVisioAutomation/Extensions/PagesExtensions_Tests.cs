using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;
using VA = VisioAutomation;
using System.Linq;
using System.Collections.Generic;
using IVisio= Microsoft.Office.Interop.Visio;

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

            Assert.AreEqual(expected.Count,actual.Count);
            Assert.AreEqual(pages[1].NameU, actual[0].NameU);
            Assert.AreEqual(pages[2].NameU, actual[1].NameU);
            Assert.AreEqual(pages[3].NameU, actual[2].NameU);

            doc1.Close(true);
        }
    }
}