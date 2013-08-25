using Microsoft.VisualStudio.TestTools.UnitTesting;
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
            var doc = ss.Document.New(page_size.Width, page_size.Height);

            var page1 = ss.Page.Get();
            ss.Page.New(page_size, false);
            var page2 = ss.Page.Get();
            ss.Page.New(page_size, false);
            var page3 = ss.Page.Get();

            Assert.AreEqual(3,doc.Pages.Count);
            Assert.AreEqual(page3, ss.Page.Get());
            ss.Page.GoTo(VA.Scripting.PageDirection.First);
            Assert.AreEqual(page1, ss.Page.Get());
            ss.Page.GoTo(VA.Scripting.PageDirection.Last);
            Assert.AreEqual(page3, ss.Page.Get());
            ss.Page.GoTo(VA.Scripting.PageDirection.Previous);
            Assert.AreEqual(page2, ss.Page.Get());
            ss.Page.GoTo(VA.Scripting.PageDirection.Next);
            Assert.AreEqual(page3, ss.Page.Get());

            // move to last and try to go to next page
            ss.Page.GoTo(VA.Scripting.PageDirection.Last);
            Assert.AreEqual(page3, ss.Page.Get());
            ss.Page.GoTo(VA.Scripting.PageDirection.Next);
            Assert.AreEqual(page3, ss.Page.Get());

            // move to first and try to go to previous page
            ss.Page.GoTo(VA.Scripting.PageDirection.First);
            Assert.AreEqual(page1, ss.Page.Get());
            ss.Page.GoTo(VA.Scripting.PageDirection.Previous);
            Assert.AreEqual(page1, ss.Page.Get());

            ss.Document.Close(true);
        }

        [TestMethod]
        public void Scripting_Page_Duplication()
        {
            var page_size = new VA.Drawing.Size(8.5, 11);
            var ss = GetScriptingSession();
            var doc = ss.Document.New(page_size.Width, page_size.Height);
            ss.Draw.Rectangle(0, 0, 1, 1);
            ss.Page.Duplicate();
        }

        [TestMethod]
        public void Scripting_Page_DuplicationToDoc()
        {
            var ss = GetScriptingSession();
            var docto = ss.Document.New();
            var docfrom = ss.Document.New();
            ss.Draw.Rectangle(0, 0, 1, 1);
            ss.Page.Duplicate(docto);
        }

    }
}