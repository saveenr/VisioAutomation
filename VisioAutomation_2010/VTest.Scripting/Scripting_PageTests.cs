using MUT=Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;
using VA=VisioAutomation;

namespace VTest.Scripting
{
    [MUT.TestClass]
    public class Scripting_PageTests : Framework.VTest
    {
        [MUT.TestMethod]
        public void Scripting_Page_NewPage()
        {
            var page_size = new VisioAutomation.Core.Size(8.5, 11);
            var client = this.GetScriptingClient();
            var doc = client.Document.NewDocument();

            client.Page.NewPage(VisioScripting.TargetDocument.Auto, new VA.Core.Size(4, 5), false);
        }


        [MUT.TestMethod]
        public void Scripting_Page_Navigation()
        {
            var page_size = new VisioAutomation.Core.Size(8.5, 11);
            var client = this.GetScriptingClient();
            var doc = client.Document.NewDocument(page_size);

            var page1 = client.Page.GetActivePage();
            client.Page.NewPage(VisioScripting.TargetDocument.Auto, page_size, false);
            var page2 = client.Page.GetActivePage();
            client.Page.NewPage(VisioScripting.TargetDocument.Auto, page_size, false);
            var page3 = client.Page.GetActivePage();


            MUT.Assert.AreEqual(3,doc.Pages.Count);
            MUT.Assert.AreEqual(page3, client.Page.GetActivePage());
            client.Page.SetActivePage(VisioScripting.TargetDocument.Auto, VisioScripting.Models.PageRelativePosition.First);
            MUT.Assert.AreEqual(page1, client.Page.GetActivePage());
            client.Page.SetActivePage(VisioScripting.TargetDocument.Auto, VisioScripting.Models.PageRelativePosition.Last);
            MUT.Assert.AreEqual(page3, client.Page.GetActivePage());
            client.Page.SetActivePage(VisioScripting.TargetDocument.Auto, VisioScripting.Models.PageRelativePosition.Previous);
            MUT.Assert.AreEqual(page2, client.Page.GetActivePage());
            client.Page.SetActivePage(VisioScripting.TargetDocument.Auto, VisioScripting.Models.PageRelativePosition.Next);
            MUT.Assert.AreEqual(page3, client.Page.GetActivePage());

            // move to last and try to go to next page
            client.Page.SetActivePage(VisioScripting.TargetDocument.Auto, VisioScripting.Models.PageRelativePosition.Last);
            MUT.Assert.AreEqual(page3, client.Page.GetActivePage());
            client.Page.SetActivePage(VisioScripting.TargetDocument.Auto, VisioScripting.Models.PageRelativePosition.Next);
            MUT.Assert.AreEqual(page3, client.Page.GetActivePage());

            // move to first and try to go to previous page
            client.Page.SetActivePage(VisioScripting.TargetDocument.Auto, VisioScripting.Models.PageRelativePosition.First);
            MUT.Assert.AreEqual(page1, client.Page.GetActivePage());
            client.Page.SetActivePage(VisioScripting.TargetDocument.Auto, VisioScripting.Models.PageRelativePosition.Previous);
            MUT.Assert.AreEqual(page1, client.Page.GetActivePage());

            doc.Close(true);
        }

        [MUT.TestMethod]
        public void Scripting_Page_Duplication()
        {
            var page_size = new VisioAutomation.Core.Size(8.5, 11);
            var client = this.GetScriptingClient();
            var doc = client.Document.NewDocument(page_size);


            client.Draw.DrawRectangle(VisioScripting.TargetPage.Auto, 0, 0, 1, 1);

            client.Page.DuplicatePage(VisioScripting.TargetPage.Auto);
            doc.Close(true);
        }

        [MUT.TestMethod]
        public void Scripting_Page_DuplicationToDoc1()
        {
            var client = this.GetScriptingClient();

            // First case: the source document is already the active document
            var doc_2_dest = client.Document.NewDocument();
            doc_2_dest.Pages.Add();
            doc_2_dest.Pages.Add();
            var doc_1_src = client.Document.NewDocument();

            doc_2_dest.Title = "DEST";
            doc_1_src.Title = "SRC";


            client.Draw.DrawRectangle(VisioScripting.TargetPage.Auto, 0, 0, 1, 1);

            var dupe_page = client.Page.DuplicatePageToDocument(VisioScripting.TargetPage.Auto, doc_2_dest);

            MUT.Assert.AreEqual(1, doc_1_src.Pages.Count);
            MUT.Assert.AreEqual(4, doc_2_dest.Pages.Count);
            doc_1_src.Close(true);
            doc_2_dest.Close(true);
        }

    }
}