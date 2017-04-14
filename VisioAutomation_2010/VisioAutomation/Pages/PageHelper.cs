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
           
            var writer = new VisioAutomation.ShapeSheet.Writers.SrcWriter();
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

                page_srcs.Add(ShapeSheet.SrcConstants.PrintLeftMargin);
                page_srcs.Add(ShapeSheet.SrcConstants.PrintCenterX);
                page_srcs.Add(ShapeSheet.SrcConstants.PrintCenterY);
                page_srcs.Add(ShapeSheet.SrcConstants.PrintOnPage);
                page_srcs.Add(ShapeSheet.SrcConstants.PrintBottomMargin);
                page_srcs.Add(ShapeSheet.SrcConstants.PrintRightMargin);
                page_srcs.Add(ShapeSheet.SrcConstants.PrintPagesX);
                page_srcs.Add(ShapeSheet.SrcConstants.PrintPagesY);
                page_srcs.Add(ShapeSheet.SrcConstants.PrintTopMargin);
                page_srcs.Add(ShapeSheet.SrcConstants.PrintPaperKind);
                page_srcs.Add(ShapeSheet.SrcConstants.PrintGrid);
                page_srcs.Add(ShapeSheet.SrcConstants.PrintPageOrientation);
                page_srcs.Add(ShapeSheet.SrcConstants.PrintScaleX);
                page_srcs.Add(ShapeSheet.SrcConstants.PrintScaleY);
                page_srcs.Add(ShapeSheet.SrcConstants.PrintPaperSource);

                page_srcs.Add(ShapeSheet.SrcConstants.PageDrawingScale);
                page_srcs.Add(ShapeSheet.SrcConstants.PageDrawingScaleType);
                page_srcs.Add(ShapeSheet.SrcConstants.PageDrawingSizeType);
                page_srcs.Add(ShapeSheet.SrcConstants.PageInhibitSnap);
                page_srcs.Add(ShapeSheet.SrcConstants.PageHeight);
                page_srcs.Add(ShapeSheet.SrcConstants.PageScale);
                page_srcs.Add(ShapeSheet.SrcConstants.PageWidth);
                page_srcs.Add(ShapeSheet.SrcConstants.PageShadowObliqueAngle);
                page_srcs.Add(ShapeSheet.SrcConstants.PageShadowOffsetX);
                page_srcs.Add(ShapeSheet.SrcConstants.PageShadowOffsetY);
                page_srcs.Add(ShapeSheet.SrcConstants.PageShadowScaleFactor);
                page_srcs.Add(ShapeSheet.SrcConstants.PageShadowType);
                page_srcs.Add(ShapeSheet.SrcConstants.PageUIVisibility);
                page_srcs.Add(ShapeSheet.SrcConstants.PageDrawingResizeType);

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

                page_srcs.Add(ShapeSheet.SrcConstants.PageLayoutAvenueSizeX);
                page_srcs.Add(ShapeSheet.SrcConstants.PageLayoutAvenueSizeY);
                page_srcs.Add(ShapeSheet.SrcConstants.PageLayoutBlockSizeX);
                page_srcs.Add(ShapeSheet.SrcConstants.PageLayoutBlockSizeY);
                page_srcs.Add(ShapeSheet.SrcConstants.PageLayoutControlAsInput);
                page_srcs.Add(ShapeSheet.SrcConstants.PageLayoutDynamicsOff);
                page_srcs.Add(ShapeSheet.SrcConstants.PageLayoutEnableGrid);
                page_srcs.Add(ShapeSheet.SrcConstants.PageLayoutLineAdjustFrom);
                page_srcs.Add(ShapeSheet.SrcConstants.PageLayoutLineAdjustTo);
                page_srcs.Add(ShapeSheet.SrcConstants.PageLayoutLineJumpCode);
                page_srcs.Add(ShapeSheet.SrcConstants.PageLayoutLineJumpFactorX);
                page_srcs.Add(ShapeSheet.SrcConstants.PageLayoutLineJumpFactorY);
                page_srcs.Add(ShapeSheet.SrcConstants.PageLayoutLineJumpStyle);
                page_srcs.Add(ShapeSheet.SrcConstants.PageLayoutLineRouteExt);
                page_srcs.Add(ShapeSheet.SrcConstants.PageLayoutLineToLineX);
                page_srcs.Add(ShapeSheet.SrcConstants.PageLayoutLineToLineY);
                page_srcs.Add(ShapeSheet.SrcConstants.PageLayoutLineToNodeX);
                page_srcs.Add(ShapeSheet.SrcConstants.PageLayoutLineToNodeY);
                page_srcs.Add(ShapeSheet.SrcConstants.PageLayoutLineJumpDirX);
                page_srcs.Add(ShapeSheet.SrcConstants.PageLayoutLineJumpDirY);
                page_srcs.Add(ShapeSheet.SrcConstants.PageLayoutShapeSplit);
                page_srcs.Add(ShapeSheet.SrcConstants.PageLayoutPlaceDepth);
                page_srcs.Add(ShapeSheet.SrcConstants.PageLayoutPlaceFlip);
                page_srcs.Add(ShapeSheet.SrcConstants.PageLayoutPlaceStyle);
                page_srcs.Add(ShapeSheet.SrcConstants.PageLayoutPlowCode);
                page_srcs.Add(ShapeSheet.SrcConstants.PageLayoutResizePage);
                page_srcs.Add(ShapeSheet.SrcConstants.PageLayoutRouteStyle);
                page_srcs.Add(ShapeSheet.SrcConstants.PageLayoutAvoidPageBreaks);
            }
        }

        internal static Drawing.Size GetSize(IVisio.Page page)
        {
            var query = new ShapeSheetQuery();
            var col_height = query.AddCell(ShapeSheet.SrcConstants.PageHeight,nameof(ShapeSheet.SrcConstants.PageHeight));
            var col_width = query.AddCell(ShapeSheet.SrcConstants.PageWidth,nameof(ShapeSheet.SrcConstants.PageWidth));

            var results = query.GetResults<double>(page.PageSheet);
            double height = results.Cells[col_height];
            double width = results.Cells[col_width];
            var s = new Drawing.Size(width, height);
            return s;
        }

        internal static void SetSize(IVisio.Page page, Drawing.Size size)
        {
            var writer = new VisioAutomation.ShapeSheet.Writers.SrcWriter();
            writer.SetFormula(VisioAutomation.ShapeSheet.SrcConstants.PageWidth, size.Width);
            writer.SetFormula(VisioAutomation.ShapeSheet.SrcConstants.PageHeight, size.Height);

            writer.Commit(page.PageSheet);
        }        

        public static short[] DropManyAutoConnectors(
            IVisio.Page page,
            ICollection<Drawing.Point> points)
        {

            if (points == null)
            {
                throw new System.ArgumentNullException(nameof(points));
            }

            // NOTE: DropMany will fail if you pass in zero items to drop

            var app = page.Application;
            var thing = app.ConnectorToolDataObject;
            int num_points = points.Count;
            var masters_obj_array = Enumerable.Repeat(thing, num_points).ToArray();
            var xy_array = Drawing.Point.ToDoubles(points).ToArray();

            System.Array outids_sa;

            page.DropManyU(masters_obj_array, xy_array, out outids_sa);

            short[] outids = (short[])outids_sa;
            return outids;
        }

    }
}