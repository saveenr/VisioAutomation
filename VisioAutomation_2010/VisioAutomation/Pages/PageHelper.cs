using System.Collections.Generic;
using System.Linq;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.Query;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Pages
{
    public static class PageHelper
    {
        private static List<VisioAutomation.ShapeSheet.SRC> page_srcs;

        public static void Duplicate(
            IVisio.Page src_page,
            IVisio.Page dest_page)
        {
            init_page_srcs();

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
                throw new System.ArgumentException("Destination Page cannot be Source Page");
            }


            if (src_page != app.ActivePage)
            {
                throw new System.ArgumentException("Source page must be active page.");
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

            // Get the Cells from the Source
            var query = new ShapeSheetQuery();
            int i = 0;
            foreach (var src in page_srcs)
            {
                query.AddCell(src,"Col"+i.ToString());
                i++;
            }
            var src_surface = new VisioAutomation.ShapeSheet.ShapeSheetSurface(src_page.PageSheet);
            var src_formulas = query.GetFormulas(src_surface);

            // Set the Cells on the Destination
           
            var writer = new ShapeSheetWriter();
            for (i = 0; i < page_srcs.Count; i++)
            {
                writer.SetFormula(page_srcs[i],src_formulas.Cells[i]);
            }

            var surface = new VisioAutomation.ShapeSheet.ShapeSheetSurface(dest_page.PageSheet);
            writer.Commit(surface);

            // make sure the new page looks like the old page
            dest_page.Background = src_page.Background;
            
            // then paste any contents from the first page
            if (num_src_shapes>0)
            {
                dest_page.Paste(copy_paste_flags);                
            }
        }

        private static void init_page_srcs()
        {
            if (page_srcs == null)
            {
                page_srcs = new List<SRC>();

                page_srcs.Add(ShapeSheet.SRCConstants.PageLeftMargin);
                page_srcs.Add(ShapeSheet.SRCConstants.CenterX);
                page_srcs.Add(ShapeSheet.SRCConstants.CenterY);
                page_srcs.Add(ShapeSheet.SRCConstants.OnPage);
                page_srcs.Add(ShapeSheet.SRCConstants.PageBottomMargin);
                page_srcs.Add(ShapeSheet.SRCConstants.PageRightMargin);
                page_srcs.Add(ShapeSheet.SRCConstants.PagesX);
                page_srcs.Add(ShapeSheet.SRCConstants.PagesY);
                page_srcs.Add(ShapeSheet.SRCConstants.PageTopMargin);
                page_srcs.Add(ShapeSheet.SRCConstants.PaperKind);
                page_srcs.Add(ShapeSheet.SRCConstants.PrintGrid);
                page_srcs.Add(ShapeSheet.SRCConstants.PrintPageOrientation);
                page_srcs.Add(ShapeSheet.SRCConstants.ScaleX);
                page_srcs.Add(ShapeSheet.SRCConstants.ScaleY);
                page_srcs.Add(ShapeSheet.SRCConstants.PaperSource);
                page_srcs.Add(ShapeSheet.SRCConstants.DrawingScale);
                page_srcs.Add(ShapeSheet.SRCConstants.DrawingScaleType);
                page_srcs.Add(ShapeSheet.SRCConstants.DrawingSizeType);
                page_srcs.Add(ShapeSheet.SRCConstants.InhibitSnap);
                page_srcs.Add(ShapeSheet.SRCConstants.PageHeight);
                page_srcs.Add(ShapeSheet.SRCConstants.PageScale);
                page_srcs.Add(ShapeSheet.SRCConstants.PageWidth);
                page_srcs.Add(ShapeSheet.SRCConstants.ShdwObliqueAngle);
                page_srcs.Add(ShapeSheet.SRCConstants.ShdwOffsetX);
                page_srcs.Add(ShapeSheet.SRCConstants.ShdwOffsetY);
                page_srcs.Add(ShapeSheet.SRCConstants.ShdwScaleFactor);
                page_srcs.Add(ShapeSheet.SRCConstants.ShdwType);
                page_srcs.Add(ShapeSheet.SRCConstants.UIVisibility);
                page_srcs.Add(ShapeSheet.SRCConstants.XGridDensity);
                page_srcs.Add(ShapeSheet.SRCConstants.XGridOrigin);
                page_srcs.Add(ShapeSheet.SRCConstants.XGridSpacing);
                page_srcs.Add(ShapeSheet.SRCConstants.XRulerDensity);
                page_srcs.Add(ShapeSheet.SRCConstants.XRulerOrigin);
                page_srcs.Add(ShapeSheet.SRCConstants.YGridDensity);
                page_srcs.Add(ShapeSheet.SRCConstants.YGridOrigin);
                page_srcs.Add(ShapeSheet.SRCConstants.YGridSpacing);
                page_srcs.Add(ShapeSheet.SRCConstants.YRulerDensity);
                page_srcs.Add(ShapeSheet.SRCConstants.YRulerOrigin);
                page_srcs.Add(ShapeSheet.SRCConstants.AvenueSizeX);
                page_srcs.Add(ShapeSheet.SRCConstants.AvenueSizeY);
                page_srcs.Add(ShapeSheet.SRCConstants.BlockSizeX);
                page_srcs.Add(ShapeSheet.SRCConstants.BlockSizeY);
                page_srcs.Add(ShapeSheet.SRCConstants.CtrlAsInput);
                page_srcs.Add(ShapeSheet.SRCConstants.DynamicsOff);
                page_srcs.Add(ShapeSheet.SRCConstants.EnableGrid);
                page_srcs.Add(ShapeSheet.SRCConstants.LineAdjustFrom);
                page_srcs.Add(ShapeSheet.SRCConstants.LineAdjustTo);
                page_srcs.Add(ShapeSheet.SRCConstants.LineJumpCode);
                page_srcs.Add(ShapeSheet.SRCConstants.LineJumpFactorX);
                page_srcs.Add(ShapeSheet.SRCConstants.LineJumpFactorY);
                page_srcs.Add(ShapeSheet.SRCConstants.LineJumpStyle);
                page_srcs.Add(ShapeSheet.SRCConstants.LineRouteExt);
                page_srcs.Add(ShapeSheet.SRCConstants.LineToLineX);
                page_srcs.Add(ShapeSheet.SRCConstants.LineToLineY);
                page_srcs.Add(ShapeSheet.SRCConstants.LineToNodeX);
                page_srcs.Add(ShapeSheet.SRCConstants.LineToNodeY);
                page_srcs.Add(ShapeSheet.SRCConstants.PageLineJumpDirX);
                page_srcs.Add(ShapeSheet.SRCConstants.PageLineJumpDirY);
                page_srcs.Add(ShapeSheet.SRCConstants.PageShapeSplit);
                page_srcs.Add(ShapeSheet.SRCConstants.PlaceDepth);
                page_srcs.Add(ShapeSheet.SRCConstants.PlaceFlip);
                page_srcs.Add(ShapeSheet.SRCConstants.PlaceStyle);
                page_srcs.Add(ShapeSheet.SRCConstants.PlowCode);
                page_srcs.Add(ShapeSheet.SRCConstants.ResizePage);
                page_srcs.Add(ShapeSheet.SRCConstants.RouteStyle);
                page_srcs.Add(ShapeSheet.SRCConstants.AvoidPageBreaks);
                page_srcs.Add(ShapeSheet.SRCConstants.DrawingResizeType);
            }
        }

        internal static Drawing.Size GetSize(IVisio.Page page)
        {
            var query = new ShapeSheetQuery();
            var col_height = query.AddCell(ShapeSheet.SRCConstants.PageHeight,"PageHeight");
            var col_width = query.AddCell(ShapeSheet.SRCConstants.PageWidth,"PageWidth");

            var page_surface = new ShapeSheetSurface(page.PageSheet);
            var results = query.GetResults<double>(page_surface);
            double height = results.Cells[col_height];
            double width = results.Cells[col_width];
            var s = new Drawing.Size(width, height);
            return s;
        }

        internal static void SetSize(IVisio.Page page, Drawing.Size size)
        {
            var writer = new ShapeSheetWriter();
            writer.SetFormula(VisioAutomation.ShapeSheet.SRCConstants.PageWidth, size.Width);
            writer.SetFormula(VisioAutomation.ShapeSheet.SRCConstants.PageHeight, size.Height);

            var surface = new VisioAutomation.ShapeSheet.ShapeSheetSurface(page.PageSheet);
            writer.Commit(surface);
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