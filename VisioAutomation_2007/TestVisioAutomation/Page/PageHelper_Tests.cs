using Microsoft.VisualStudio.TestTools.UnitTesting;
using VA = VisioAutomation;

namespace TestVisioAutomation
{
    [TestClass]
    public class PageHelper_Tests : VisioAutomationTest
    {
        [TestMethod]
        public void Page_Query()
        {
            var page1 = GetNewPage(new VA.Drawing.Size(4, 3));
            var pagecells = VA.Pages.PageCells.GetCells(page1.PageSheet);
            Assert.AreEqual(new VA.Drawing.Size(4, 3), new VA.Drawing.Size(pagecells.PageWidth.Result,pagecells.PageHeight.Result));

            // Double each side
            pagecells.PageWidth = pagecells.PageWidth.Result * 2.0;
            pagecells.PageHeight = pagecells.PageHeight.Result * 2.0;

            var update = new VA.ShapeSheet.Update();
            update.SetFormulas(pagecells);
            update.Execute(page1.PageSheet);

            var pagecells2 = VA.Pages.PageCells.GetCells(page1.PageSheet);
            Assert.AreEqual(new VA.Drawing.Size(8, 6), new VA.Drawing.Size(pagecells2.PageWidth.Result, pagecells2.PageHeight.Result));
            page1.Delete(0);
        }

        [TestMethod]
        public void Page_Orientation()
        {
            var page1 = GetNewPage(new VA.Drawing.Size(4, 3));

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