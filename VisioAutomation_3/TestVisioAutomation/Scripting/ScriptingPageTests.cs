using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation;
using VAS = VisioAutomation.Scripting;
using VA = VisioAutomation;

namespace TestVisioAutomation
{
    [TestClass]
    public class ScriptingPageTests : VisioAutomationTest
    {
        [TestMethod]
        public void Scripting_Page_Navigation()
        {
            var page_size = new VA.Drawing.Size(8.5, 11);
            var ss = GetScriptingSession();
            var doc = ss.Document.NewDocument(page_size.Width, page_size.Height);

            var page1 = ss.Page.GetPage();
            ss.Page.NewPage(page_size, false);
            var page2 = ss.Page.GetPage();
            ss.Page.NewPage(page_size, false);
            var page3 = ss.Page.GetPage();

            Assert.AreEqual(3,doc.Pages.Count);
            Assert.AreEqual(page3, ss.Page.GetPage());
            ss.Page.NavigateToPage(PageNavigation.FirstPage);
            Assert.AreEqual(page1, ss.Page.GetPage());
            ss.Page.NavigateToPage(PageNavigation.LastPage);
            Assert.AreEqual(page3, ss.Page.GetPage());
            ss.Page.NavigateToPage(PageNavigation.PreviousPage);
            Assert.AreEqual(page2, ss.Page.GetPage());
            ss.Page.NavigateToPage(PageNavigation.NextPage);
            Assert.AreEqual(page3, ss.Page.GetPage());

            // move to last and try to go to next page
            ss.Page.NavigateToPage(PageNavigation.LastPage);
            Assert.AreEqual(page3, ss.Page.GetPage());
            ss.Page.NavigateToPage(PageNavigation.NextPage);
            Assert.AreEqual(page3, ss.Page.GetPage());

            // move to first and try to go to previous page
            ss.Page.NavigateToPage(PageNavigation.FirstPage);
            Assert.AreEqual(page1, ss.Page.GetPage());
            ss.Page.NavigateToPage(PageNavigation.PreviousPage);
            Assert.AreEqual(page1, ss.Page.GetPage());

            ss.Document.CloseDocument(true);
        }
    }
}