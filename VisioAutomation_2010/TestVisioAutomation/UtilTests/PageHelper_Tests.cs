using Microsoft.VisualStudio.TestTools.UnitTesting;
using VA = VisioAutomation;

namespace TestVisioAutomation
{
    [TestClass]
    public class PageHelper_Tests : VisioAutomationTest
    {
        [TestMethod]
        public void PageScenarios()
        {
            this.PageOrientation();
            this.DuplicatePage();
        }

        public void PageOrientation()
        {
            var page1 = GetNewPage(new VA.Drawing.Size(4, 3));

            var ss = this.GetScriptingSession();

            var or1 = ss.Page.GetOrientation();
            Assert.AreEqual(VA.Pages.PrintPageOrientation.Portrait, or1);

            var size1 = ss.Page.GetSize();
            Assert.AreEqual(new VA.Drawing.Size(4, 3), size1);

            ss.Page.SetOrientation(VA.Pages.PrintPageOrientation.Landscape);

            var or2 = ss.Page.GetOrientation();
            Assert.AreEqual(VA.Pages.PrintPageOrientation.Landscape, or2);
            var size2 = ss.Page.GetSize();
            Assert.AreEqual(new VA.Drawing.Size(3, 4), size2);

            page1.Delete(0);
        }

        public void DuplicatePage()
        {
            var page1 = GetNewPage(new VA.Drawing.Size(4, 3));
            var s1 = page1.DrawRectangle(1, 1, 3, 3);

            var doc = page1.Document;
            var pages = doc.Pages;

            var page2 = pages.Add();

            // Activate Page 1 - needed for duplicate to work
            var app = page1.Application;
            var active_window = app.ActiveWindow;
            active_window.Page = page1;

            VA.Pages.PageHelper.Duplicate(page1, page2);

            Assert.AreEqual(new VA.Drawing.Size(4, 3), VisioAutomationTest.GetPageSize(page2));
            Assert.AreEqual(1, page2.Shapes.Count);

            page2.Delete(0);
            page1.Delete(0);
        }
    }
}