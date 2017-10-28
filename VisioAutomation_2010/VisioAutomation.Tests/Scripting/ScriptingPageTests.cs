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
            var docto_1 = client.Document.NewDocument();
            var docfrom_1 = client.Document.NewDocument();


            var target_src_page = new VisioScripting.Models.TargetPage();
            target_src_page.Resolve(client);

            client.Draw.DrawRectangle(0, 0, 1, 1);
            client.Page.DuplicatePageToDocument(target_src_page, docto_1);

            // Second case: the source document has to be activated beforehand
            var docfrom_2 = client.Document.NewDocument();
            var docto_2 = client.Document.NewDocument();
            client.Document.ActivateDocument(docfrom_2);
            client.Draw.DrawRectangle(0, 0, 1, 1);
            client.Page.DuplicatePageToDocument(target_src_page, docto_2);

            docfrom_1.Close(true);
            docfrom_2.Close(true);
            docto_1.Close(true);
            docto_2.Close(true);
        }

        [TestMethod]
        public void PSTestGetVisioPageCell()
        {
            var client = this.GetScriptingClient();

            var visDoc = client.Document.NewDocument();
/*            var results1 = client.Page.

            var cells1 = VisioPowerShellTests.invoker.Invoke("Get-VisioPageCell -Cells PageWidth,PageHeight -Page (Get-VisioPage -ActivePage) -GetResults -ResultType Double");
            var data_row_collection1 = (DataRowCollection)cells1[0].Properties["Rows"].Value;
            var results = data_row_collection1[0];
            Assert.IsNotNull(cells1);
            Assert.AreEqual(8.5, results["PageWidth"]);
            Assert.AreEqual(11.0, results["PageHeight"]);

            //Now lets add another page and get it's width and height
            var page2 = VisioPowerShellTests.invoker.Invoke("New-VisioPage");
            var cells2 = VisioPowerShellTests.invoker.Invoke("Get-VisioPageCell -Cells PageWidth,PageHeight -Page (Get-VisioPage -ActivePage) -GetResults -ResultType Double");
            var data_row_collection2 = (DataRowCollection)cells2[0].Properties["Rows"].Value;
            var results2 = data_row_collection2[0];

            Assert.IsNotNull(cells2);
            Assert.AreEqual(8.5, results2["PageWidth"]);
            Assert.AreEqual(11.0, results2["PageHeight"]);

            // Close Visio Application that was created when "New-VisioDocument" was invoked
            VisioPowerShellTests.invoker.Invoke("Close-VisioApplication -Force");
 * 
 * */
        }
    }
}