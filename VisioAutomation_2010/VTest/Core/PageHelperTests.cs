using VisioAutomation.Extensions;
using VTest.Framework;
using MUT=Microsoft.VisualStudio.TestTools.UnitTesting;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VTest.Core
{
    [MUT.TestClass]
    public class PageHelperTests : Framework.VTest
    {
        [MUT.TestMethod]
        public void Page_Query()
        {
            var size = new VA.Core.Size(4, 3);
            var page1 = this.GetNewPage(size);
            var page_fmt_cells = VA.Pages.FormatCells.GetCells(page1.PageSheet, VisioAutomation.Core.CellValueType.Formula);
            MUT.Assert.AreEqual("4 in", page_fmt_cells.Width.Value);
            MUT.Assert.AreEqual("3 in", page_fmt_cells.Height.Value);

            // Double each side
            var page_fmt_cells1 = page_fmt_cells;
            page_fmt_cells1.Width = "8";
            page_fmt_cells1.Height = "6";

            var writer = new VisioAutomation.ShapeSheet.Writers.SrcWriter();
            writer.SetValues(page_fmt_cells1);

            writer.Commit(page1.PageSheet, VisioAutomation.Core.CellValueType.Formula);

            var actual_page_format_cells = VA.Pages.FormatCells.GetCells(page1.PageSheet, VisioAutomation.Core.CellValueType.Result);
            MUT.Assert.AreEqual("8.0000 in.", actual_page_format_cells.Width.Value);
            MUT.Assert.AreEqual("6.0000 in.", actual_page_format_cells.Height.Value);
            page1.Delete(0);
        }

        [MUT.TestMethod]
        public void Page_Orientation()
        {
            var size = new VA.Core.Size(4, 3);

            var page1 = this.GetNewPage(size);

            var client = this.GetScriptingClient();

            var orientation_1 = client.Page.GetPageOrientation(VisioScripting.TargetPage.Auto);
            MUT.Assert.AreEqual(VisioScripting.Models.PageOrientation.Portrait, orientation_1);

            var sizes_1 = client.Page.GetPageSize(VisioScripting.TargetPages.Auto);
            MUT.Assert.AreEqual(size, sizes_1[0]);

            var target_pages = new VisioScripting.TargetPages(page1);
            client.Page.SetPageOrientation(target_pages, VisioScripting.Models.PageOrientation.Landscape);

            var orientation_2 = client.Page.GetPageOrientation(VisioScripting.TargetPage.Auto);
            MUT.Assert.AreEqual(VisioScripting.Models.PageOrientation.Landscape, orientation_2);

            var actual_final_sizes = client.Page.GetPageSize(VisioScripting.TargetPages.Auto);
            var expected_final_size = new VA.Core.Size(3, 4);
            MUT.Assert.AreEqual(expected_final_size, actual_final_sizes[0]);

            page1.Delete(0);
        }

        [MUT.TestMethod]
        public void Page_Duplicate()
        {
            var page_size = new VA.Core.Size(4, 3);
            var page1 = this.GetNewPage(page_size);
            var s1 = page1.DrawRectangle(1, 1, 3, 3);

            var doc = page1.Document;
            var pages = doc.Pages;

            var page2 = pages.Add();

            // Activate Page 1 - needed for duplicate to work
            var app = page1.Application;
            var active_window = app.ActiveWindow;
            active_window.Page = page1;

            VA.Pages.PageHelper.Duplicate(page1, page2);

            MUT.Assert.AreEqual(page_size, GetPageSize(page2));
            MUT.Assert.AreEqual(1, page2.Shapes.Count);

            page2.Delete(0);
            page1.Delete(0);
        }

        [MUT.TestMethod]
        public void Page_SwitchPages()
        {
            var app = this.GetVisioApplication();

            var documents = app.Documents;
            int old_doc_count = documents.Count;

            var doc1 = this.GetNewDoc();
            MUT.Assert.AreEqual(documents.Count, old_doc_count + 1);
            MUT.Assert.AreEqual(doc1.Pages.Count, 1);
            var page1 = doc1.Pages[1];
            MUT.Assert.AreEqual(app.ActivePage, page1);

            var page2 = doc1.Pages.Add();
            page2.Background = 0;
            SetPageSize(page2, this.StandardPageSize);

            var active_window = app.ActiveWindow;
            MUT.Assert.AreEqual(app.ActivePage, page2);
            active_window.Page = page1;
            MUT.Assert.AreEqual(app.ActivePage, page1);
            active_window.Page = page2;
            MUT.Assert.AreEqual(app.ActivePage, page2);
            doc1.Close(true);
        }

        [MUT.TestMethod]
        public void Page_ResizeBorder()
        {
            var doc = this.GetNewDoc();
            var shapesize = new VisioAutomation.Core.Size(1, 2);
            var border1 = new VisioAutomation.Core.Size(0, 0);
            var border2 = new VA.Core.Size(3, 4);
            _verify_page_size_to_fit(doc, new VA.Core.Size(1, 1), new VA.Core.Size(1, 1), shapesize, border1, 1.5, 2);
            _verify_page_size_to_fit(doc, new VA.Core.Size(0, 0), new VA.Core.Size(0, 0), shapesize, border1, 0.5, 1);
            _verify_page_size_to_fit(doc, new VA.Core.Size(1, 0), new VA.Core.Size(0, 0), shapesize, border1, 1.5, 1);
            _verify_page_size_to_fit(doc, new VA.Core.Size(0, 1), new VA.Core.Size(0, 0), shapesize, border1, 0.5, 2);
            _verify_page_size_to_fit(doc, new VA.Core.Size(0, 0), new VA.Core.Size(1, 0), shapesize, border1, 0.5, 1);
            _verify_page_size_to_fit(doc, new VA.Core.Size(0, 0), new VA.Core.Size(0, 1), shapesize, border1, 0.5, 1);
            _verify_page_size_to_fit(doc, new VA.Core.Size(1, 1), new VA.Core.Size(1, 1), shapesize, border2, 4.5, 6);
            _verify_page_size_to_fit(doc, new VA.Core.Size(1, 0), new VA.Core.Size(0, 0), shapesize, border2, 4, 5);
            _verify_page_size_to_fit(doc, new VA.Core.Size(0, 1), new VA.Core.Size(0, 0), shapesize, border2, 3.5, 5.5);
            _verify_page_size_to_fit(doc, new VA.Core.Size(0, 0), new VA.Core.Size(1, 0), shapesize, border2, 4, 5);
            _verify_page_size_to_fit(doc, new VA.Core.Size(0, 0), new VA.Core.Size(0, 1), shapesize, border2, 3.5, 5.5);
            doc.Close(true);
        }

        private static void _verify_page_size_to_fit(IVisio.Document doc,
            VA.Core.Size bottomleft_margin,
            VA.Core.Size upperright_margin,
            VA.Core.Size shape_size,
            VA.Core.Size padding_size,
            double expected_pinx,
            double expected_piny)
        {
            var page = doc.Pages.Add();

            var pagecells = new VA.Pages.PrintCells();
            pagecells.TopMargin = upperright_margin.Height;
            pagecells.BottomMargin = bottomleft_margin.Height;
            pagecells.LeftMargin = bottomleft_margin.Width;
            pagecells.RightMargin = upperright_margin.Width;

            var page_writer = new VisioAutomation.ShapeSheet.Writers.SrcWriter();
            page_writer.SetValues(pagecells);

            page_writer.Commit(page.PageSheet, VisioAutomation.Core.CellValueType.Formula);


            var shape = page.DrawRectangle(5, 5, 5 + shape_size.Width, 5 + shape_size.Height);
            page.ResizeToFitContents(padding_size);
            var xform = VA.Shapes.XFormCells.GetCells(shape, VisioAutomation.Core.CellValueType.Result);
            var pinpos = xform.GetPinPosResult();
            MUT.Assert.AreEqual(expected_pinx, pinpos.X, 0.1);
            MUT.Assert.AreEqual(expected_piny, pinpos.Y, 0.1);
            page.Delete(0);
        }
    }
}