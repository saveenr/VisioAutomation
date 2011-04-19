using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Extensions;
using IVisio=Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;

namespace VisioAutomation
{
    public static class PageHelper
    {
        public static void SetBackgroundPage(IVisio.Page fgpage, IVisio.Page bgpage)
        {
            if (fgpage == null)
            {
                throw new System.ArgumentNullException("fgpage");
            }

            if (bgpage != null)
            {
                // Set the background page
                // Check that the intended background is indeed a background page
                if (bgpage.Background == 0)
                {
                    string msg = string.Format("Page \"{0}\" is not a background page", bgpage.Name);
                    throw new AutomationException(msg);
                }

                // don't allow the page to be set as a background to itself
                if (fgpage == bgpage)
                {
                    string msg = string.Format("Cannot set page as its own background page");
                    throw new AutomationException(msg);
                }

                // Finally set it
                fgpage.BackPage = bgpage;
            }
            else
            {
                // Clear the background page
                fgpage.BackPage = string.Empty;
            }
        }

        private static void copy_page_cells(IVisio.Page src_page, IVisio.Page dest_page)
        {
            if (src_page == null)
            {
                throw new System.ArgumentNullException("src_page");
            }

            if (dest_page == null)
            {
                throw new System.ArgumentNullException("dest_page");
            }

            var src_pagesheet = src_page.PageSheet;
            var dest_pagesheet = dest_page.PageSheet;

            // Collect the cells from the source page
            var pagecells = VA.Layout.PageCells.GetCells(src_pagesheet);

            // Set them on the destination page
            var update = new VA.ShapeSheet.Update.SRCUpdate();
            pagecells.Apply(update);
            update.Execute(dest_pagesheet);
        }

        public static VA.Layout.PageCells GetPageCells(IVisio.Page page)
        {
            var pagesheet = page.PageSheet;
            var pagecells = VA.Layout.PageCells.GetCells(pagesheet);
            return pagecells;
        }

        public static IVisio.Page Duplicate(IVisio.Page page)
        {
            if (page == null)
            {
                throw new System.ArgumentNullException("page");
            }

            var application = page.Application;
            if (page != application.ActivePage)
            {
                throw new System.ArgumentException("Source page must be active page.", "page");
            }

            bool do_copy = true;
            bool has_something_to_copy = false;
            IVisio.Shape shape_to_copy = null;
            bool perform_group_before_copy = false;
            var copy_flag = IVisio.VisCutCopyPasteCodes.visCopyPasteNoTranslate;

            int num_shapes = 0;
            if (do_copy && (page.Shapes.Count > 0))
            {
                has_something_to_copy = true;
                var active_window = application.ActiveWindow;
                active_window.SelectAll();
                var selection = active_window.Selection;
                num_shapes = selection.Count;

                if (num_shapes == 1)
                {
                    shape_to_copy = page.Shapes[1];
                    perform_group_before_copy = false;
                }
                else
                {
                    var w = application.ActiveWindow;
                    var sel = w.Selection;
                    shape_to_copy = sel.Group();
                    perform_group_before_copy = true;
                }

                shape_to_copy.Copy(copy_flag);

                if (perform_group_before_copy)
                {
                    shape_to_copy.Ungroup();
                }
            }

            var new_page = page.Document.Pages.Add();
            new_page.Background = 0; // ensure this is a foreground page
            new_page.SetSize(page.GetSize()); // set the size

            copy_page_cells(page, new_page);
            if (do_copy && has_something_to_copy)
            {
                new_page.Paste(copy_flag);

                if (perform_group_before_copy)
                {
                    var active_window = application.ActiveWindow;
                    var selection = active_window.Selection;
                    selection.Ungroup();
                }
            }

            return new_page;
        }

        public static void DuplicateToDocument(
            IVisio.Page src_page,
            IVisio.Document dest_doc,
            IVisio.Page dest_page,
            string new_page_name,
            bool suppress_ui)
        {
            if (src_page == null)
            {
                throw new System.ArgumentNullException("src_page");
            }

            if (dest_page == null)
            {
                throw new System.ArgumentNullException("dest_page");
            }

            if (dest_doc == null)
            {
                throw new System.ArgumentNullException("dest_doc");
            }

            // http://support.microsoft.com/kb/290581
            var app = src_page.Application;

            short copy_paste_flags = (short)IVisio.VisCutCopyPasteCodes.visCopyPasteNoTranslate;

            src_page.Document.Activate();
            src_page.Activate();

            bool has_content = src_page.Shapes.Count > 0;

            // copy contents
            if (has_content)
            {
                src_page.Activate();
                var active_window = app.ActiveWindow;
                active_window.SelectAll();
                var selection = active_window.Selection;
                selection.Copy(copy_paste_flags);
                active_window.DeselectAll();
            }

            // Create the new page and give it the same general properties
            copy_page_cells(src_page, dest_page);

            // Try to preserve the name
            new_page_name = new_page_name ?? src_page.Name;
            var existing_names = new HashSet<string>(dest_doc.Pages.GetNamesU());
            if (!existing_names.Contains(new_page_name))
            {
                dest_page.Name = new_page_name;
            }

            // paste any contents 
            if (has_content)
            {
                using (var alertresponse = app.CreateAlertResponseScope(VA.UI.AlertResponseCode.Ignore))
                {
                    dest_page.Paste(copy_paste_flags);
                }
            }
        }

        public static VA.Layout.PrintPageOrientation GetOrientation(IVisio.Page page)
        {
            if (page == null)
            {
                throw new System.ArgumentNullException("page");
            }

            var page_sheet = page.PageSheet;
            var orientationcell = page_sheet.GetCell(VA.ShapeSheet.SRCConstants.PrintPageOrientation);
            int value = orientationcell.ResultInt[IVisio.VisUnitCodes.visNoCast, 0];
            return (VA.Layout.PrintPageOrientation)value;
        }

        /// <summary>
        /// Sets the orientation of the page
        /// </summary>
        /// <param name="page"></param>
        /// <param name="orientation">1=portrait, 2=landscape</param>
        public static void SetOrientation(IVisio.Page page, VA.Layout.PrintPageOrientation orientation)
        {
            if (page == null)
            {
                throw new System.ArgumentNullException("page");
            }

            if (orientation != VA.Layout.PrintPageOrientation.Landscape && orientation != VA.Layout.PrintPageOrientation.Portrait)
            {
                throw new System.ArgumentOutOfRangeException("orientation", "must be either Portrait or Landscape");
            }

            var old_orientation = GetOrientation(page);

            if (old_orientation == orientation)
            {
                // don't need to do anything
                return;
            }

            var srcs = new[] { VA.ShapeSheet.SRCConstants.PageWidth, VA.ShapeSheet.SRCConstants.PageHeight };
            var stream = VA.ShapeSheet.SRCStream.FromItems(srcs);
            var unitcodes = new[] { IVisio.VisUnitCodes.visNoCast, IVisio.VisUnitCodes.visNoCast };
            var page_sheet = page.PageSheet;
            var results = VA.ShapeSheet.ShapeSheetHelper.GetResults<double>(page_sheet, stream, unitcodes);

            double old_width = results[0];
            double old_height = results[1];

            double new_height = old_width;
            double new_width = old_height;

            var update = new VA.ShapeSheet.Update.SRCUpdate(3, 0);
            update.SetFormula(VA.ShapeSheet.SRCConstants.PageWidth, new_width);
            update.SetFormula(VA.ShapeSheet.SRCConstants.PageHeight, new_height);
            update.SetFormula(VA.ShapeSheet.SRCConstants.PrintPageOrientation, (int)orientation);

            update.Execute(page_sheet);
        }

        public static VA.Drawing.Size GetSize(IVisio.Page page)
        {
            if (page == null)
            {
                throw new System.ArgumentNullException("page");
            }

            var srcs = new[] { VA.ShapeSheet.SRCConstants.PageWidth, VA.ShapeSheet.SRCConstants.PageHeight };
            var stream = VA.ShapeSheet.SRCStream.FromItems(srcs);
            var unitcodes = new[] { IVisio.VisUnitCodes.visNoCast, IVisio.VisUnitCodes.visNoCast };
            var pagesheet = page.PageSheet;
            var results = VA.ShapeSheet.ShapeSheetHelper.GetResults<double>(pagesheet, stream, unitcodes);
            var s = new VA.Drawing.Size(results[0], results[1]);
            return s;
        }

        public static void SetSize(IVisio.Page page, VA.Drawing.Size size)
        {
            if (page == null)
            {
                throw new System.ArgumentNullException("page");
            }

            var page_sheet = page.PageSheet;

            var update = new VA.ShapeSheet.Update.SRCUpdate(2, 0);
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
                page.Activate();
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

        public static void Activate(IVisio.Page page)
        {
            var app = page.Application;
            // If the page belongs to an Inactive Document, then activate that document first
            if (app.ActiveDocument != page.Document)
            {
                var page_doc = page.Document;
                page_doc.Activate();
            }

            // Double-check: make sure the page's parent document is active
            if (app.ActiveDocument != page.Document)
            {
                var page_doc = page.Document;
                string msg = string.Format("Failed to activate document \"{0}\"", page_doc.Name);
                throw new AutomationException(msg);
            }

            var active_window = app.ActiveWindow;
            // if current page is not already active, then activate it
            if (active_window.Page != page)
            {
                active_window.Page = page;
            }

            // Double-check: make sure the page is active now
            if (active_window.Page != page)
            {
                string msg = string.Format("Failed to activate page \"{0}\"", page.Name);
                throw new AutomationException(msg);
            }
        }

        public static void ResizeToFitContents(IVisio.Page page, VA.Drawing.Size bordersize)
        {
            page.ResizeToFitContents();

            if ((bordersize.Width > 0.0) || (bordersize.Height > 0.0))
            {
                var old_size = page.GetSize();
                var new_size = old_size + bordersize.Multiply(2, 2);
                page.SetSize(new_size);
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
            var xy_array = VA.Drawing.DrawingUtil.PointsToDoubles(points).ToArray();

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

            var update = new VA.ShapeSheet.Update.SRCUpdate();

            update.SetFormula(VA.ShapeSheet.SRCConstants.XGridOrigin, "0.0");
            update.SetFormula(VA.ShapeSheet.SRCConstants.YGridOrigin, "0.0");
            update.SetFormula(VA.ShapeSheet.SRCConstants.XRulerOrigin, "0.0");
            update.SetFormula(VA.ShapeSheet.SRCConstants.YRulerOrigin, "0.0");

            update.Execute(page.PageSheet);
        }
    }
}