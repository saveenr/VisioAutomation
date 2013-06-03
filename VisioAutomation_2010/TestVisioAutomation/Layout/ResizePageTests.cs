using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace TestVisioAutomation
{
    [TestClass]
    public class ResizePageTests : VisioAutomationTest
    {
        [TestMethod]
        public void ResizeBorder()
        {
            var doc = this.GetNewDoc();
            var shapesize = new VA.Drawing.Size(1, 2);
            var border1 = new VA.Drawing.Size(0, 0);
            var border2 = new VA.Drawing.Size(3, 4);
            TestResize(doc, new VA.Drawing.Size(1, 1), new VA.Drawing.Size(1, 1), shapesize, border1, 1.5, 2);
            TestResize(doc, new VA.Drawing.Size(0, 0), new VA.Drawing.Size(0, 0), shapesize, border1, 0.5, 1);
            TestResize(doc, new VA.Drawing.Size(1, 0), new VA.Drawing.Size(0, 0), shapesize, border1, 1.5, 1);
            TestResize(doc, new VA.Drawing.Size(0, 1), new VA.Drawing.Size(0, 0), shapesize, border1, 0.5, 2);
            TestResize(doc, new VA.Drawing.Size(0, 0), new VA.Drawing.Size(1, 0), shapesize, border1, 0.5, 1);
            TestResize(doc, new VA.Drawing.Size(0, 0), new VA.Drawing.Size(0, 1), shapesize, border1, 0.5, 1);
            TestResize(doc, new VA.Drawing.Size(1, 1), new VA.Drawing.Size(1, 1), shapesize, border2, 4.5, 6);
            TestResize(doc, new VA.Drawing.Size(1, 0), new VA.Drawing.Size(0, 0), shapesize, border2, 4, 5);
            TestResize(doc, new VA.Drawing.Size(0, 1), new VA.Drawing.Size(0, 0), shapesize, border2, 3.5, 5.5);
            TestResize(doc, new VA.Drawing.Size(0, 0), new VA.Drawing.Size(1, 0), shapesize, border2, 4, 5);
            TestResize(doc, new VA.Drawing.Size(0, 0), new VA.Drawing.Size(0, 1), shapesize, border2, 3.5, 5.5);
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

            var pageupdate = new VA.ShapeSheet.Update();
            pageupdate.SetFormulas(pagecells);
            pageupdate.Execute(page.PageSheet);


            var shape = page.DrawRectangle(5, 5, 5 + shape_size.Width, 5+shape_size.Height);
            page.ResizeToFitContents(padding_size);
            var xform = VA.Layout.XFormCells.GetCells(shape);
            AssertVA.AreEqual(expected_pinx, expected_piny, xform.Pin(), 0.1);
            page.Delete(0);
        }
    }
}