using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;
using VisioAutomation.Geometry;
using VisioScripting.Models;

namespace VisioAutomation_Tests.Scripting
{
    [TestClass]
    public class ScriptingPageTests : VisioAutomationTest
    {
        [TestMethod]
        public void Scripting_Page_NewPage()
        {
            var page_size = new VisioAutomation.Geometry.Size(8.5, 11);
            var client = this.GetScriptingClient();
            var doc = client.Document.NewDocument();
            client.Page.NewPage(new Size(4, 5), false);
        }


        [TestMethod]
        public void Scripting_Page_Navigation()
        {
            var page_size = new VisioAutomation.Geometry.Size(8.5, 11);
            var client = this.GetScriptingClient();
            var doc = client.Document.NewDocument(page_size);

            var page1 = client.Page.GetActivePage();
            client.Page.NewPage(page_size, false);
            var page2 = client.Page.GetActivePage();
            client.Page.NewPage(page_size, false);
            var page3 = client.Page.GetActivePage();

            Assert.AreEqual(3,doc.Pages.Count);
            Assert.AreEqual(page3, client.Page.GetActivePage());
            client.Page.SetActivePageByDirection(PageDirection.First);
            Assert.AreEqual(page1, client.Page.GetActivePage());
            client.Page.SetActivePageByDirection(PageDirection.Last);
            Assert.AreEqual(page3, client.Page.GetActivePage());
            client.Page.SetActivePageByDirection(PageDirection.Previous);
            Assert.AreEqual(page2, client.Page.GetActivePage());
            client.Page.SetActivePageByDirection(PageDirection.Next);
            Assert.AreEqual(page3, client.Page.GetActivePage());

            // move to last and try to go to next page
            client.Page.SetActivePageByDirection(PageDirection.Last);
            Assert.AreEqual(page3, client.Page.GetActivePage());
            client.Page.SetActivePageByDirection(PageDirection.Next);
            Assert.AreEqual(page3, client.Page.GetActivePage());

            // move to first and try to go to previous page
            client.Page.SetActivePageByDirection(PageDirection.First);
            Assert.AreEqual(page1, client.Page.GetActivePage());
            client.Page.SetActivePageByDirection(PageDirection.Previous);
            Assert.AreEqual(page1, client.Page.GetActivePage());

            doc.Close(true);
        }

        [TestMethod]
        public void Scripting_Page_Duplication()
        {
            var page_size = new VisioAutomation.Geometry.Size(8.5, 11);
            var client = this.GetScriptingClient();
            var doc = client.Document.NewDocument(page_size);
            client.Draw.DrawRectangle(0, 0, 1, 1);
            client.Page.DuplicateActivePage();
            doc.Close(true);
        }

        [TestMethod]
        public void Scripting_Page_DuplicationToDoc1()
        {
            var client = this.GetScriptingClient();

            // First case: the source document is already the active document
            var doc_dest_1 = client.Document.NewDocument();
            var doc_src_1 = client.Document.NewDocument();


            var target_src1_page = new VisioScripting.TargetPage().Resolve(client);

            client.Draw.DrawRectangle(0, 0, 1, 1);
            client.Page.DuplicatePageToDocument(target_src1_page, doc_dest_1);

            // Second case: the source document has to be activated beforehand
            var doc_src_2 = client.Document.NewDocument();
            var doc_dest_2 = client.Document.NewDocument();
            client.Document.ActivateDocument(doc_src_2);
            client.Draw.DrawRectangle(0, 0, 1, 1);

            var target_src2_page = new VisioScripting.TargetPage().Resolve(client);

            client.Page.DuplicatePageToDocument(target_src2_page, doc_dest_2);

            doc_src_1.Close(true);
            doc_src_2.Close(true);
            doc_dest_1.Close(true);
            doc_dest_2.Close(true);
        }
    }
}