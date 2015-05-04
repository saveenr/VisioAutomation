using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Drawing;
using VisioAutomation.Extensions;
using VisioAutomation.Pages;
using VisioAutomation.Shapes;
using VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace TestVisioAutomation
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
            var shapesize = new Size(1, 2);
            var border1 = new Size(0, 0);
            var border2 = new Size(3, 4);
            ResizePageTests.TestResize(doc, new Size(1, 1), new Size(1, 1), shapesize, border1, 1.5, 2);
            ResizePageTests.TestResize(doc, new Size(0, 0), new Size(0, 0), shapesize, border1, 0.5, 1);
            ResizePageTests.TestResize(doc, new Size(1, 0), new Size(0, 0), shapesize, border1, 1.5, 1);
            ResizePageTests.TestResize(doc, new Size(0, 1), new Size(0, 0), shapesize, border1, 0.5, 2);
            ResizePageTests.TestResize(doc, new Size(0, 0), new Size(1, 0), shapesize, border1, 0.5, 1);
            ResizePageTests.TestResize(doc, new Size(0, 0), new Size(0, 1), shapesize, border1, 0.5, 1);
            ResizePageTests.TestResize(doc, new Size(1, 1), new Size(1, 1), shapesize, border2, 4.5, 6);
            ResizePageTests.TestResize(doc, new Size(1, 0), new Size(0, 0), shapesize, border2, 4, 5);
            ResizePageTests.TestResize(doc, new Size(0, 1), new Size(0, 0), shapesize, border2, 3.5, 5.5);
            ResizePageTests.TestResize(doc, new Size(0, 0), new Size(1, 0), shapesize, border2, 4, 5);
            ResizePageTests.TestResize(doc, new Size(0, 0), new Size(0, 1), shapesize, border2, 3.5, 5.5);
            doc.Close(true);
        }
        
        private static void TestResize(IVisio.Document doc, 
            Size bottomleft_margin, 
            Size upperright_margin, 
            Size shape_size, 
            Size padding_size,  
            double expected_pinx, 
            double expected_piny)
        {
            var page = doc.Pages.Add();

            var pagecells = new PageCells();
            pagecells.PageTopMargin = upperright_margin.Height;
            pagecells.PageBottomMargin = bottomleft_margin.Height;
            pagecells.PageLeftMargin = bottomleft_margin.Width;
            pagecells.PageRightMargin = upperright_margin.Width;

            var pageupdate = new Update();
            pageupdate.SetFormulas(pagecells);
            pageupdate.Execute(page.PageSheet);


            var shape = page.DrawRectangle(5, 5, 5 + shape_size.Width, 5+shape_size.Height);
            page.ResizeToFitContents(padding_size);
            var xform = XFormCells.GetCells(shape);
            AssertVA.AreEqual(expected_pinx, expected_piny, xform.Pin(), 0.1);
            page.Delete(0);
        }
    }
}