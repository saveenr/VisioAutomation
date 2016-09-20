using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.ShapeSheet.Writers;
using VA = VisioAutomation;

namespace VisioAutomation_Tests.Core.Page
{
    [TestClass]
    public class PageHelperTests : VisioAutomationTest
    {
        [TestMethod]
        public void Page_Query()
        {
            var page1 = this.GetNewPage(new VA.Drawing.Size(4, 3));
            var pagecells = VA.Pages.PageCells.GetCells(page1.PageSheet);
            Assert.AreEqual("4.0000 in.", pagecells.PageWidth.Result);
            Assert.AreEqual("3.0000 in.", pagecells.PageHeight.Result);

            // Double each side
            pagecells.PageWidth = "8.0";
            pagecells.PageHeight = "6.0";

            var writer = new FormulaWriterSRC();
            pagecells.SetFormulas(writer);
            writer.Commit(page1.PageSheet);

            var pagecells2 = VA.Pages.PageCells.GetCells(page1.PageSheet);
            Assert.AreEqual("8.0000 in.", pagecells2.PageWidth.Result);
            Assert.AreEqual("6.0000 in.", pagecells2.PageHeight.Result);
            page1.Delete(0);
        }

        [TestMethod]
        public void Page_Orientation()
        {
            var page1 = this.GetNewPage(new VA.Drawing.Size(4, 3));

            var client = this.GetScriptingClient();

            var or1 = client.Page.GetOrientation();
            Assert.AreEqual(VA.Pages.PrintPageOrientation.Portrait, or1);

            var size1 = client.Page.GetSize();
            Assert.AreEqual(new VA.Drawing.Size(4, 3), size1);

            client.Page.SetOrientation(VA.Pages.PrintPageOrientation.Landscape);

            var or2 = client.Page.GetOrientation();
            Assert.AreEqual(VA.Pages.PrintPageOrientation.Landscape, or2);
            var size2 = client.Page.GetSize();
            Assert.AreEqual(new VA.Drawing.Size(3, 4), size2);

            page1.Delete(0);
        }

        [TestMethod]
        public void Page_Duplicate()
        {
            var page1 = this.GetNewPage(new VA.Drawing.Size(4, 3));
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