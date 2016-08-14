using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;
using VisioAutomation.ShapeSheet.Writers;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation_Tests.Core
{
    [TestClass]
    public class ResizePageTests : VisioAutomationTest
    {
        [TestMethod]
        public void Page_SwitchPages()
        {
            var app = this.GetVisioApplication();

            var documents = app.Documents;
            int old_doc_count = documents.Count;

            var doc1 = this.GetNewDoc();
            Assert.AreEqual(documents.Count, old_doc_count + 1);
            Assert.AreEqual(doc1.Pages.Count, 1);
            var page1 = doc1.Pages[1];
            Assert.AreEqual(app.ActivePage, page1);

            var page2 = doc1.Pages.Add();
            page2.Background = 0;
            VisioAutomationTest.SetPageSize(page2, this.StandardPageSize);

            var active_window = app.ActiveWindow;
            Assert.AreEqual(app.ActivePage, page2);
            active_window.Page = page1;
            Assert.AreEqual(app.ActivePage, page1);
            active_window.Page = page2;
            Assert.AreEqual(app.ActivePage, page2);
            doc1.Close(true);
        }

        [TestMethod]
        public void Page_ResizeBorder()
        {
            var doc = this.GetNewDoc();
            var shapesize = new VisioAutomation.Drawing.Size(1, 2);
            var border1 = new VisioAutomation.Drawing.Size(0, 0);
            var border2 = new VA.Drawing.Size(3, 4);
            ResizePageTests.TestResize(doc, new VA.Drawing.Size(1, 1), new VA.Drawing.Size(1, 1), shapesize, border1, 1.5, 2);
            ResizePageTests.TestResize(doc, new VA.Drawing.Size(0, 0), new VA.Drawing.Size(0, 0), shapesize, border1, 0.5, 1);
            ResizePageTests.TestResize(doc, new VA.Drawing.Size(1, 0), new VA.Drawing.Size(0, 0), shapesize, border1, 1.5, 1);
            ResizePageTests.TestResize(doc, new VA.Drawing.Size(0, 1), new VA.Drawing.Size(0, 0), shapesize, border1, 0.5, 2);
            ResizePageTests.TestResize(doc, new VA.Drawing.Size(0, 0), new VA.Drawing.Size(1, 0), shapesize, border1, 0.5, 1);
            ResizePageTests.TestResize(doc, new VA.Drawing.Size(0, 0), new VA.Drawing.Size(0, 1), shapesize, border1, 0.5, 1);
            ResizePageTests.TestResize(doc, new VA.Drawing.Size(1, 1), new VA.Drawing.Size(1, 1), shapesize, border2, 4.5, 6);
            ResizePageTests.TestResize(doc, new VA.Drawing.Size(1, 0), new VA.Drawing.Size(0, 0), shapesize, border2, 4, 5);
            ResizePageTests.TestResize(doc, new VA.Drawing.Size(0, 1), new VA.Drawing.Size(0, 0), shapesize, border2, 3.5, 5.5);
            ResizePageTests.TestResize(doc, new VA.Drawing.Size(0, 0), new VA.Drawing.Size(1, 0), shapesize, border2, 4, 5);
            ResizePageTests.TestResize(doc, new VA.Drawing.Size(0, 0), new VA.Drawing.Size(0, 1), shapesize, border2, 3.5, 5.5);
            doc.Close(true);
        }
        
        private static void TestResize(IVisio.Document doc,
            VA.Drawing.Size bottomleft_margin,
            VA.Drawing.Size upperright_margin,
            VA.Drawing.Size shape_size,
            VA.Drawing.Size padding_size,  
            double expected_pinx, 
            double expected_piny)
        {
            var page = doc.Pages.Add();

            var pagecells = new VA.Pages.PageCells();
            pagecells.PageTopMargin = upperright_margin.Height;
            pagecells.PageBottomMargin = bottomleft_margin.Height;
            pagecells.PageLeftMargin = bottomleft_margin.Width;
            pagecells.PageRightMargin = upperright_margin.Width;

            var pageupdate = new FormulaWriterSRC();
            pagecells.SetFormulas(pageupdate);
            pageupdate.Commit(page.PageSheet);


            var shape = page.DrawRectangle(5, 5, 5 + shape_size.Width, 5+shape_size.Height);
            page.ResizeToFitContents(padding_size);
            var xform = VA.Shapes.XFormCells.GetCells(shape);
            AssertUtil.AreEqual(expected_pinx, expected_piny, xform.GetPinPosResult(), 0.1);
            page.Delete(0);
        }
    }
}