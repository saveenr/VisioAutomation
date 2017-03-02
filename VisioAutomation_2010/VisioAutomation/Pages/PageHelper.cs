using System.Collections.Generic;
using System.Linq;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.Query;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Pages
{
    public static class PageHelper
    {
        private static List<VisioAutomation.ShapeSheet.Src> page_srcs;

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

            var src_formulas = query.GetFormulas(src_page.PageSheet);

            // Set the Cells on the Destination
           
            var writer = new ShapeSheetWriter();
            for (i = 0; i < page_srcs.Count; i++)
            {
                writer.SetFormula(page_srcs[i],src_formulas.Cells[i]);
            }

            writer.Commit(dest_page.PageSheet);

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
                page_srcs = new List<Src>();

                page_srcs.Add(ShapeSheet.SrcConstants.PageLeftMargin);
                page_srcs.Add(ShapeSheet.SrcConstants.CenterX);
                page_srcs.Add(ShapeSheet.SrcConstants.CenterY);
                page_srcs.Add(ShapeSheet.SrcConstants.OnPage);
                page_srcs.Add(ShapeSheet.SrcConstants.PageBottomMargin);
                page_srcs.Add(ShapeSheet.SrcConstants.PageRightMargin);
                page_srcs.Add(ShapeSheet.SrcConstants.PagesX);
                page_srcs.Add(ShapeSheet.SrcConstants.PagesY);
                page_srcs.Add(ShapeSheet.SrcConstants.PageTopMargin);
                page_srcs.Add(ShapeSheet.SrcConstants.PaperKind);
                page_srcs.Add(ShapeSheet.SrcConstants.PrintGrid);
                page_srcs.Add(ShapeSheet.SrcConstants.PrintPageOrientation);
                page_srcs.Add(ShapeSheet.SrcConstants.ScaleX);
                page_srcs.Add(ShapeSheet.SrcConstants.ScaleY);
                page_srcs.Add(ShapeSheet.SrcConstants.PaperSource);
                page_srcs.Add(ShapeSheet.SrcConstants.DrawingScale);
                page_srcs.Add(ShapeSheet.SrcConstants.DrawingScaleType);
                page_srcs.Add(ShapeSheet.SrcConstants.DrawingSizeType);
                page_srcs.Add(ShapeSheet.SrcConstants.InhibitSnap);
                page_srcs.Add(ShapeSheet.SrcConstants.PageHeight);
                page_srcs.Add(ShapeSheet.SrcConstants.PageScale);
                page_srcs.Add(ShapeSheet.SrcConstants.PageWidth);
                page_srcs.Add(ShapeSheet.SrcConstants.ShdwObliqueAngle);
                page_srcs.Add(ShapeSheet.SrcConstants.ShdwOffsetX);
                page_srcs.Add(ShapeSheet.SrcConstants.ShdwOffsetY);
                page_srcs.Add(ShapeSheet.SrcConstants.ShdwScaleFactor);
                page_srcs.Add(ShapeSheet.SrcConstants.ShdwType);
                page_srcs.Add(ShapeSheet.SrcConstants.UIVisibility);
                page_srcs.Add(ShapeSheet.SrcConstants.XGridDensity);
                page_srcs.Add(ShapeSheet.SrcConstants.XGridOrigin);
                page_srcs.Add(ShapeSheet.SrcConstants.XGridSpacing);
                page_srcs.Add(ShapeSheet.SrcConstants.XRulerDensity);
                page_srcs.Add(ShapeSheet.SrcConstants.XRulerOrigin);
                page_srcs.Add(ShapeSheet.SrcConstants.YGridDensity);
                page_srcs.Add(ShapeSheet.SrcConstants.YGridOrigin);
                page_srcs.Add(ShapeSheet.SrcConstants.YGridSpacing);
                page_srcs.Add(ShapeSheet.SrcConstants.YRulerDensity);
                page_srcs.Add(ShapeSheet.SrcConstants.YRulerOrigin);
                page_srcs.Add(ShapeSheet.SrcConstants.AvenueSizeX);
                page_srcs.Add(ShapeSheet.SrcConstants.AvenueSizeY);
                page_srcs.Add(ShapeSheet.SrcConstants.BlockSizeX);
                page_srcs.Add(ShapeSheet.SrcConstants.BlockSizeY);
                page_srcs.Add(ShapeSheet.SrcConstants.CtrlAsInput);
                page_srcs.Add(ShapeSheet.SrcConstants.DynamicsOff);
                page_srcs.Add(ShapeSheet.SrcConstants.EnableGrid);
                page_srcs.Add(ShapeSheet.SrcConstants.LineAdjustFrom);
                page_srcs.Add(ShapeSheet.SrcConstants.LineAdjustTo);
                page_srcs.Add(ShapeSheet.SrcConstants.LineJumpCode);
                page_srcs.Add(ShapeSheet.SrcConstants.LineJumpFactorX);
                page_srcs.Add(ShapeSheet.SrcConstants.LineJumpFactorY);
                page_srcs.Add(ShapeSheet.SrcConstants.LineJumpStyle);
                page_srcs.Add(ShapeSheet.SrcConstants.LineRouteExt);
                page_srcs.Add(ShapeSheet.SrcConstants.LineToLineX);
                page_srcs.Add(ShapeSheet.SrcConstants.LineToLineY);
                page_srcs.Add(ShapeSheet.SrcConstants.LineToNodeX);
                page_srcs.Add(ShapeSheet.SrcConstants.LineToNodeY);
                page_srcs.Add(ShapeSheet.SrcConstants.PageLineJumpDirX);
                page_srcs.Add(ShapeSheet.SrcConstants.PageLineJumpDirY);
                page_srcs.Add(ShapeSheet.SrcConstants.PageShapeSplit);
                page_srcs.Add(ShapeSheet.SrcConstants.PlaceDepth);
                page_srcs.Add(ShapeSheet.SrcConstants.PlaceFlip);
                page_srcs.Add(ShapeSheet.SrcConstants.PlaceStyle);
                page_srcs.Add(ShapeSheet.SrcConstants.PlowCode);
                page_srcs.Add(ShapeSheet.SrcConstants.ResizePage);
                page_srcs.Add(ShapeSheet.SrcConstants.RouteStyle);
                page_srcs.Add(ShapeSheet.SrcConstants.AvoidPageBreaks);
                page_srcs.Add(ShapeSheet.SrcConstants.DrawingResizeType);
            }
        }

        internal static Drawing.Size GetSize(IVisio.Page page)
        {
            var query = new ShapeSheetQuery();
            var col_height = query.AddCell(ShapeSheet.SrcConstants.PageHeight,"PageHeight");
            var col_width = query.AddCell(ShapeSheet.SrcConstants.PageWidth,"PageWidth");

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
            writer.SetFormula(VisioAutomation.ShapeSheet.SrcConstants.PageWidth, size.Width);
            writer.SetFormula(VisioAutomation.ShapeSheet.SrcConstants.PageHeight, size.Height);

            writer.Commit(page.PageSheet);
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