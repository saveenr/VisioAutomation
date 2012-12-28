using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Extensions;
using IVisio=Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;

namespace VisioAutomation.Pages
{
    public static class PageHelper
    {
        public static VA.Pages.PageCells GetPageCells(IVisio.Page page)
        {
            var pagesheet = page.PageSheet;
            var pagecells = VA.Pages.PageCells.GetCells(pagesheet);
            return pagecells;
        }

        public static IVisio.Page Duplicate(IVisio.Page src_page,string dest_page_name)
        {
            var pcc = new VA.Internal.PageContentCopier(src_page);

            var doc = src_page.Document;
            var pages = doc.Pages;
            var dest_page = pages.Add();
            dest_page.Name = dest_page_name;

            pcc.ApplyTo(dest_page);

            return dest_page;
        }

        
        public static void Duplicate(
            IVisio.Page src_page,
            IVisio.Page dest_page,
            string dest_page_name)
        {
            var app = src_page.Application;
            var doc = src_page.Document;
            dest_page.Name = dest_page_name;

            var pcc = new VA.Internal.PageContentCopier(src_page);
            pcc.ApplyTo(dest_page);
        }

        public static VA.Pages.PrintPageOrientation GetOrientation(IVisio.Page page)
        {
            if (page == null)
            {
                throw new System.ArgumentNullException("page");
            }

            var page_sheet = page.PageSheet;
            var src = VA.ShapeSheet.SRCConstants.PrintPageOrientation;
            var orientationcell = page_sheet.CellsSRC[src.Section, src.Row, src.Cell];
            int value = orientationcell.ResultInt[IVisio.VisUnitCodes.visNoCast, 0];
            return (VA.Pages.PrintPageOrientation)value;
        }

        /// <summary>
        /// Sets the orientation of the page
        /// </summary>
        /// <param name="page"></param>
        /// <param name="orientation">1=portrait, 2=landscape</param>
        public static void SetOrientation(IVisio.Page page, VA.Pages.PrintPageOrientation orientation)
        {
            if (page == null)
            {
                throw new System.ArgumentNullException("page");
            }

            if (orientation != VA.Pages.PrintPageOrientation.Landscape && orientation != VA.Pages.PrintPageOrientation.Portrait)
            {
                throw new System.ArgumentOutOfRangeException("orientation", "must be either Portrait or Landscape");
            }

            var old_orientation = GetOrientation(page);

            if (old_orientation == orientation)
            {
                // don't need to do anything
                return;
            }

            var old_size = VA.Pages.PageHelper.GetSize(page);

            double new_height = old_size.Width;
            double new_width = old_size.Height;

            var update = new VA.ShapeSheet.Update(3);
            update.SetFormula(VA.ShapeSheet.SRCConstants.PageWidth, new_width);
            update.SetFormula(VA.ShapeSheet.SRCConstants.PageHeight, new_height);
            update.SetFormula(VA.ShapeSheet.SRCConstants.PrintPageOrientation, (int)orientation);

            update.Execute(page.PageSheet);
        }

        public static VA.Drawing.Size GetSize(IVisio.Page page)
        {
            if (page == null)
            {
                throw new System.ArgumentNullException("page");
            }

            var query = new VA.ShapeSheet.Query.CellQuery();
            var col_height = query.AddColumn(VA.ShapeSheet.SRCConstants.PageHeight);
            var col_width = query.AddColumn(VA.ShapeSheet.SRCConstants.PageWidth);
            var results = query.GetResults<double>(page.PageSheet);
            double height = results[0, col_height];
            double width = results[0, col_width];
            var s = new VA.Drawing.Size(width, height);
            return s;
        }

        public static void SetSize(IVisio.Page page, VA.Drawing.Size size)
        {
            if (page == null)
            {
                throw new System.ArgumentNullException("page");
            }

            var page_sheet = page.PageSheet;

            var update = new VA.ShapeSheet.Update(2);
            update.SetFormula(VA.ShapeSheet.SRCConstants.PageWidth, size.Width);
            update.SetFormula(VA.ShapeSheet.SRCConstants.PageHeight, size.Height);
            update.Execute(page_sheet);
        }

        public static void NavigateTo(IVisio.Pages pages, PageNavigation flags)
        {
            if (pages == null)
            {
                throw new System.ArgumentNullException("pages");
            }

            var app = pages.Application;
            var active_document = app.ActiveDocument;
            if (pages.Document != active_document)
            {
                throw new System.ArgumentException("Page.Document is not application's ActiveDocument");
            }

            if (pages.Count < 2)
            {
                throw new AutomationException("Only 1 page available. Navigation not possible.");
            }

            var activepage = app.ActivePage;

            int cur_index = activepage.Index;
            const int min_index = 1;
            int max_index = pages.Count;
            int new_index = move_in_range(cur_index, min_index, max_index, flags);
            if (cur_index != new_index)
            {
                var doc_pages = active_document.Pages;
                var page = doc_pages[new_index];

                var active_window = app.ActiveWindow;
                active_window.Page = page;
            }
        }

        internal static int move_in_range(int cur, int min, int max, PageNavigation direction)
        {
            if (max < min)
            {
                throw new System.ArgumentOutOfRangeException("max");
            }

            if (cur < min)
            {
                throw new System.ArgumentOutOfRangeException("cur");
            }

            if (cur > max)
            {
                throw new System.ArgumentOutOfRangeException("cur");
            }

            if (direction == PageNavigation.NextPage)
            {
                return System.Math.Min(cur + 1, max);
            }
            else if (direction == PageNavigation.PreviousPage)
            {
                return System.Math.Max(cur - 1, min);
            }
            else if (direction == PageNavigation.FirstPage)
            {
                return min;
            }
            else if (direction == PageNavigation.LastPage)
            {
                return max;
            }
            else
            {
                throw new System.ArgumentOutOfRangeException("direction");
            }
        }

        public static void ResizeToFitContents(IVisio.Page page, VA.Drawing.Size bordersize)
        {
            page.ResizeToFitContents();

            if ((bordersize.Width > 0.0) || (bordersize.Height > 0.0))
            {
                var old_size = VA.Pages.PageHelper.GetSize(page);
                var new_size = old_size + bordersize.Multiply(2, 2);
                VA.Pages.PageHelper.SetSize(page,new_size);
                page.CenterDrawing();
            }
        }

        public static short[] DropManyU(
            IVisio.Page page,
            IList<IVisio.Master> masters,
            IEnumerable<VA.Drawing.Point> points)
        {
            if (masters == null)
            {
                throw new System.ArgumentNullException("masters");
            }

            if (masters.Count < 1)
            {
                return new short[0];
            }

            if (points == null)
            {
                throw new System.ArgumentNullException("points");
            }

            // NOTE: DropMany will fail if you pass in zero items to drop
            var masters_obj_array = masters.Cast<object>().ToArray();
            var xy_array = VA.Drawing.Point.ToDoubles(points).ToArray();

            System.Array outids_sa;

            page.DropManyU(masters_obj_array, xy_array, out outids_sa);

            short[] outids = (short[])outids_sa;
            return outids;
        }

        public static void ResetOrigin(IVisio.Page page)
        {
            if (page == null)
            {
                throw new System.ArgumentNullException("page");
            }

            var update = new VA.ShapeSheet.Update();

            update.SetFormula(VA.ShapeSheet.SRCConstants.XGridOrigin, "0.0");
            update.SetFormula(VA.ShapeSheet.SRCConstants.YGridOrigin, "0.0");
            update.SetFormula(VA.ShapeSheet.SRCConstants.XRulerOrigin, "0.0");
            update.SetFormula(VA.ShapeSheet.SRCConstants.YRulerOrigin, "0.0");

            update.Execute(page.PageSheet);
        }
    }
}