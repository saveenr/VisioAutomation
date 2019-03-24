using System.Collections.Generic;
using System.Linq;
using VASS=VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Pages
{
    public static class PageHelper
    {
        private static List<VASS.Src> page_srcs;

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
            var query = new VASS.Query.CellQuery();
            int i = 0;
            foreach (var src in page_srcs)
            {
                query.Columns.Add(src,"Col"+i.ToString());
                i++;
            }

            var src_formulas = query.GetFormulas(src_page.PageSheet);

            // Set the Cells on the Destination
           
            var writer = new VASS.Writers.SrcWriter();
            for (i = 0; i < page_srcs.Count; i++)
            {
                int row = 0;
                writer.SetValue(page_srcs[i],src_formulas[row][i]);
            }

            writer.CommitFormulas(dest_page.PageSheet);

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
                page_srcs = new List<VASS.Src>();

                page_srcs.Add(VASS.SrcConstants.PrintLeftMargin);
                page_srcs.Add(VASS.SrcConstants.PrintCenterX);
                page_srcs.Add(VASS.SrcConstants.PrintCenterY);
                page_srcs.Add(VASS.SrcConstants.PrintOnPage);
                page_srcs.Add(VASS.SrcConstants.PrintBottomMargin);
                page_srcs.Add(VASS.SrcConstants.PrintRightMargin);
                page_srcs.Add(VASS.SrcConstants.PrintPagesX);
                page_srcs.Add(VASS.SrcConstants.PrintPagesY);
                page_srcs.Add(VASS.SrcConstants.PrintTopMargin);
                page_srcs.Add(VASS.SrcConstants.PrintPaperKind);
                page_srcs.Add(VASS.SrcConstants.PrintGrid);
                page_srcs.Add(VASS.SrcConstants.PrintPageOrientation);
                page_srcs.Add(VASS.SrcConstants.PrintScaleX);
                page_srcs.Add(VASS.SrcConstants.PrintScaleY);
                page_srcs.Add(VASS.SrcConstants.PrintPaperSource);

                page_srcs.Add(VASS.SrcConstants.PageDrawingScale);
                page_srcs.Add(VASS.SrcConstants.PageDrawingScaleType);
                page_srcs.Add(VASS.SrcConstants.PageDrawingSizeType);
                page_srcs.Add(VASS.SrcConstants.PageInhibitSnap);
                page_srcs.Add(VASS.SrcConstants.PageHeight);
                page_srcs.Add(VASS.SrcConstants.PageScale);
                page_srcs.Add(VASS.SrcConstants.PageWidth);
                page_srcs.Add(VASS.SrcConstants.PageShadowObliqueAngle);
                page_srcs.Add(VASS.SrcConstants.PageShadowOffsetX);
                page_srcs.Add(VASS.SrcConstants.PageShadowOffsetY);
                page_srcs.Add(VASS.SrcConstants.PageShadowScaleFactor);
                page_srcs.Add(VASS.SrcConstants.PageShadowType);
                page_srcs.Add(VASS.SrcConstants.PageUIVisibility);
                page_srcs.Add(VASS.SrcConstants.PageDrawingResizeType);

                page_srcs.Add(VASS.SrcConstants.XGridDensity);
                page_srcs.Add(VASS.SrcConstants.XGridOrigin);
                page_srcs.Add(VASS.SrcConstants.XGridSpacing);
                page_srcs.Add(VASS.SrcConstants.XRulerDensity);
                page_srcs.Add(VASS.SrcConstants.XRulerOrigin);
                page_srcs.Add(VASS.SrcConstants.YGridDensity);
                page_srcs.Add(VASS.SrcConstants.YGridOrigin);
                page_srcs.Add(VASS.SrcConstants.YGridSpacing);
                page_srcs.Add(VASS.SrcConstants.YRulerDensity);
                page_srcs.Add(VASS.SrcConstants.YRulerOrigin);

                page_srcs.Add(VASS.SrcConstants.PageLayoutAvenueSizeX);
                page_srcs.Add(VASS.SrcConstants.PageLayoutAvenueSizeY);
                page_srcs.Add(VASS.SrcConstants.PageLayoutBlockSizeX);
                page_srcs.Add(VASS.SrcConstants.PageLayoutBlockSizeY);
                page_srcs.Add(VASS.SrcConstants.PageLayoutControlAsInput);
                page_srcs.Add(VASS.SrcConstants.PageLayoutDynamicsOff);
                page_srcs.Add(VASS.SrcConstants.PageLayoutEnableGrid);
                page_srcs.Add(VASS.SrcConstants.PageLayoutLineAdjustFrom);
                page_srcs.Add(VASS.SrcConstants.PageLayoutLineAdjustTo);
                page_srcs.Add(VASS.SrcConstants.PageLayoutLineJumpCode);
                page_srcs.Add(VASS.SrcConstants.PageLayoutLineJumpFactorX);
                page_srcs.Add(VASS.SrcConstants.PageLayoutLineJumpFactorY);
                page_srcs.Add(VASS.SrcConstants.PageLayoutLineJumpStyle);
                page_srcs.Add(VASS.SrcConstants.PageLayoutLineRouteExt);
                page_srcs.Add(VASS.SrcConstants.PageLayoutLineToLineX);
                page_srcs.Add(VASS.SrcConstants.PageLayoutLineToLineY);
                page_srcs.Add(VASS.SrcConstants.PageLayoutLineToNodeX);
                page_srcs.Add(VASS.SrcConstants.PageLayoutLineToNodeY);
                page_srcs.Add(VASS.SrcConstants.PageLayoutLineJumpDirX);
                page_srcs.Add(VASS.SrcConstants.PageLayoutLineJumpDirY);
                page_srcs.Add(VASS.SrcConstants.PageLayoutShapeSplit);
                page_srcs.Add(VASS.SrcConstants.PageLayoutPlaceDepth);
                page_srcs.Add(VASS.SrcConstants.PageLayoutPlaceFlip);
                page_srcs.Add(VASS.SrcConstants.PageLayoutPlaceStyle);
                page_srcs.Add(VASS.SrcConstants.PageLayoutPlowCode);
                page_srcs.Add(VASS.SrcConstants.PageLayoutResizePage);
                page_srcs.Add(VASS.SrcConstants.PageLayoutRouteStyle);
                page_srcs.Add(VASS.SrcConstants.PageLayoutAvoidPageBreaks);
            }
        }

        internal static Geometry.Size GetSize(IVisio.Page page)
        {
            var query = new VASS.Query.CellQuery();
            var col_height = query.Columns.Add(VASS.SrcConstants.PageHeight,nameof(VASS.SrcConstants.PageHeight));
            var col_width = query.Columns.Add(VASS.SrcConstants.PageWidth,nameof(VASS.SrcConstants.PageWidth));

            var cellqueryresult = query.GetResults<double>(page.PageSheet);
            var row = cellqueryresult[0];
            double height = row[col_height];
            double width = row[col_width];
            var s = new Geometry.Size(width, height);
            return s;
        }

        internal static void SetSize(IVisio.Page page, Geometry.Size size)
        {
            var writer = new VASS.Writers.SrcWriter();
            writer.SetValue(VASS.SrcConstants.PageWidth, size.Width);
            writer.SetValue(VASS.SrcConstants.PageHeight, size.Height);

            writer.CommitFormulas(page.PageSheet);
        }        

        public static short[] DropManyAutoConnectors(
            IVisio.Page page,
            ICollection<Geometry.Point> points)
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
            var xy_array = Geometry.Point.ToDoubles(points).ToArray();

            System.Array outids_sa;

            page.DropManyU(masters_obj_array, xy_array, out outids_sa);

            short[] outids = (short[])outids_sa;
            return outids;
        }
    }
}