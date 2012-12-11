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

            var page_size1 = new VA.Drawing.Size(10, 5);
            var page_size2 = new VA.Drawing.Size(6, 3);

            page1.SetSize(page_size1);
            Assert.AreEqual(page_size1, page1.GetSize());

            page1.SetSize(page_size2);
            Assert.AreEqual(page_size2, page1.GetSize());
            page1.Delete(0);
        }

        [TestMethod]
        public void TestAsEnumerable()
        {
            var doc1 = this.GetNewDoc();
            var page1 = doc1.Pages[1];
            var page2 = doc1.Pages.Add();
            var page3 = doc1.Pages.Add();

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