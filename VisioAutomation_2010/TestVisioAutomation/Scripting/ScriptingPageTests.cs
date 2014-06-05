using Microsoft.VisualStudio.TestTools.UnitTesting;
using VA = VisioAutomation;
using VisioAutomation.Extensions;

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

            doc.Close(true);
        }

        [TestMethod]
        public void Scripting_Page_Duplication()
        {
            var page_size = new VA.Drawing.Size(8.5, 11);
            var ss = GetScriptingSession();
            var doc = ss.Document.New(page_size.Width, page_size.Height);
            ss.Draw.Rectangle(0, 0, 1, 1);
            ss.Page.Duplicate();
            doc.Close(true);
        }

        [TestMethod]
        public void Scripting_Page_DuplicationToDoc1()
        {
            var ss = GetScriptingSession();

            // First case: the source document is already the active document
            var docto_1 = ss.Document.New();
            var docfrom_1 = ss.Document.New();
            ss.Draw.Rectangle(0, 0, 1, 1);
            ss.Page.Duplicate(docto_1);

            // Second case: the source document has to be activated beforehand
            var docfrom_2 = ss.Document.New();
            var docto_2 = ss.Document.New();
            VA.Documents.DocumentHelper.Activate(docfrom_2);
            ss.Draw.Rectangle(0, 0, 1, 1);
            ss.Page.Duplicate(docto_2);

            docfrom_1.Close(true);
            docfrom_2.Close(true);
            docto_1.Close(true);
            docto_2.Close(true);
        }
    }
}