using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Pages
{
    public static class PageHelper
    {
        public static void Duplicate(
            IVisio.Page src_page,
            IVisio.Page dest_page)
        {
            var app = src_page.Application;
            short copy_paste_flags = (short)IVisio.VisCutCopyPasteCodes.visCopyPasteNoTranslate;

            // handle the source page
            if (src_page == null)
            {
                throw new System.ArgumentNullException(nameof(src_page));
            }

            if (dest_page == null)
            {
                throw new System.ArgumentNullException(nameof(dest_page));
            }

            if (dest_page == src_page)
            {
                throw new AutomationException("Destination Page cannot be Source Page");
            }


            if (src_page != app.ActivePage)
            {
                throw new AutomationException("Source page must be active page.");
            }

            var src_page_shapes = src_page.Shapes;
            int num_src_shapes=src_page_shapes.Count;

            if (num_src_shapes > 0)
            {
                var active_window = app.ActiveWindow;
                active_window.SelectAll();
                var selection = active_window.Selection;
                selection.Copy(copy_paste_flags);
                active_window.DeselectAll();
            }

            var src_pagesheet = src_page.PageSheet;
            var pagecells = PageCells.GetCells(src_pagesheet);


            // handle the dest page

            // first update all the page cells
            var dest_pagesheet = dest_page.PageSheet;
            var update = new ShapeSheet.Update();
            update.SetFormulas(pagecells);
            update.Execute(dest_pagesheet);

            // make sure the new page looks like the old page
            dest_page.Background = src_page.Background;
            
            // then paste any contents from the first page
            if (num_src_shapes>0)
            {
                dest_page.Paste(copy_paste_flags);                
            }
        }

        private static Drawing.Size GetSize(IVisio.Page page)
        {
            var query = new ShapeSheet.Query.CellQuery();
            var col_height = query.AddCell(ShapeSheet.SRCConstants.PageHeight,"PageHeight");
            var col_width = query.AddCell(ShapeSheet.SRCConstants.PageWidth,"PageWidth");
            var results = query.GetResults<double>(page.PageSheet);
            double height = results[col_height];
            double width = results[col_width];
            var s = new Drawing.Size(width, height);
            return s;
        }

        private static void SetSize(IVisio.Page page, Drawing.Size size)
        {
            var page_cells = new PageCells();
            page_cells.PageHeight = size.Height;
            page_cells.PageWidth = size.Width;
            var pageupdate = new ShapeSheet.Update();
            pageupdate.SetFormulas(page_cells);
            pageupdate.Execute(page.PageSheet);
        }
        
        public static void ResizeToFitContents(IVisio.Page page, Drawing.Size padding)
        {
            // first perform the native resizetofit
            page.ResizeToFitContents();

            if ((padding.Width > 0.0) || (padding.Height > 0.0))
            {
                // if there is any additional padding requested
                // we need to further handle the page

                // first determine the desired page size including the padding
                // and set the new size

                var old_size = PageHelper.GetSize(page);
                var new_size = old_size + padding.Multiply(2, 2);
                PageHelper.SetSize(page,new_size);

                // The page has the correct size, but
                // the contents will be offset from the correct location
                page.CenterDrawing();
            }
        }

        public static short[] DropManyU(
            IVisio.Page page,
            IList<IVisio.Master> masters,
            IEnumerable<Drawing.Point> points)
        {
            if (masters == null)
            {
                throw new System.ArgumentNullException(nameof(masters));
            }

            if (masters.Count < 1)
            {
                return new short[0];
            }

            if (points == null)
            {
                throw new System.ArgumentNullException(nameof(points));
            }

            // NOTE: DropMany will fail if you pass in zero items to drop
            var masters_obj_array = masters.Cast<object>().ToArray();
            var xy_array = Drawing.Point.ToDoubles(points).ToArray();

            System.Array outids_sa;

            page.DropManyU(masters_obj_array, xy_array, out outids_sa);

            short[] outids = (short[])outids_sa;
            return outids;
        }

        public static short[] DropManyAutoConnectors(
            IVisio.Page page,
            IEnumerable<Drawing.Point> points)
        {

            if (points == null)
            {
                throw new System.ArgumentNullException(nameof(points));
            }

            // NOTE: DropMany will fail if you pass in zero items to drop

            var app = page.Application;
            var thing = app.ConnectorToolDataObject;
            int num_points = points.Count();
            var masters_obj_array = Enumerable.Repeat(thing, num_points).ToArray();
            var xy_array = Drawing.Point.ToDoubles(points).ToArray();

            System.Array outids_sa;

            page.DropManyU(masters_obj_array, xy_array, out outids_sa);

            short[] outids = (short[])outids_sa;
            return outids;
        }

    }
}