using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;

using VA = VisioAutomation;

namespace TestVisioAutomation
{
    [TestClass]
    public class ScriptingPageTests : VisioAutomationTest
    {
        [TestMethod]
        public void Scripting_Page_Navigation()
        {
            var page_size = new VisioAutomation.Drawing.Size(8.5, 11);
            var client = this.GetScriptingClient();
            var doc = client.Document.New(page_size.Width, page_size.Height);

            var page1 = client.Page.Get();
            client.Page.New(page_size, false);
            var page2 = client.Page.Get();
            client.Page.New(page_size, false);
            var page3 = client.Page.Get();

            Assert.AreEqual(3,doc.Pages.Count);
            Assert.AreEqual(page3, client.Page.Get());
            client.Page.GoTo(VisioAutomation.Scripting.PageDirection.First);
            Assert.AreEqual(page1, client.Page.Get());
            client.Page.GoTo(VisioAutomation.Scripting.PageDirection.Last);
            Assert.AreEqual(page3, client.Page.Get());
            client.Page.GoTo(VisioAutomation.Scripting.PageDirection.Previous);
            Assert.AreEqual(page2, client.Page.Get());
            client.Page.GoTo(VisioAutomation.Scripting.PageDirection.Next);
            Assert.AreEqual(page3, client.Page.Get());

            // move to last and try to go to next page
            client.Page.GoTo(VisioAutomation.Scripting.PageDirection.Last);
            Assert.AreEqual(page3, client.Page.Get());
            client.Page.GoTo(VisioAutomation.Scripting.PageDirection.Next);
            Assert.AreEqual(page3, client.Page.Get());

            // move to first and try to go to previous page
            client.Page.GoTo(VisioAutomation.Scripting.PageDirection.First);
            Assert.AreEqual(page1, client.Page.Get());
            client.Page.GoTo(VisioAutomation.Scripting.PageDirection.Previous);
            Assert.AreEqual(page1, client.Page.Get());

            doc.Close(true);
        }

        [TestMethod]
        public void Scripting_Page_Duplication()
        {
            var page_size = new VisioAutomation.Drawing.Size(8.5, 11);
            var client = this.GetScriptingClient();
            var doc = client.Document.New(page_size.Width, page_size.Height);
            client.Draw.Rectangle(0, 0, 1, 1);
            client.Page.Duplicate();
            doc.Close(true);
        }

        [TestMethod]
        public void Scripting_Page_DuplicationToDoc1()
        {
            var client = this.GetScriptingClient();

            // First case: the source document is already the active document
            var docto_1 = client.Document.New();
            var docfrom_1 = client.Document.New();
            client.Draw.Rectangle(0, 0, 1, 1);
            client.Page.Duplicate(docto_1);

            // Second case: the source document has to be activated beforehand
            var docfrom_2 = client.Document.New();
            var docto_2 = client.Document.New();
            VisioAutomation.Documents.DocumentHelper.Activate(docfrom_2);
            client.Draw.Rectangle(0, 0, 1, 1);
            client.Page.Duplicate(docto_2);

            docfrom_1.Close(true);
            docfrom_2.Close(true);
            docto_1.Close(true);
            docto_2.Close(true);
        }
    }
}