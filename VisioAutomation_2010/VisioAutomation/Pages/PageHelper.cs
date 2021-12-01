using System.Collections.Generic;
using System.Linq;
using VASS = VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Pages
{
    public static class PageHelper
    {
        private static List<Core.Src> _static_page_srcs;

        public static void Duplicate(
            IVisio.Page src_page,
            IVisio.Page dest_page)
        {
            init_page_srcs();

            var app = src_page.Application;
            short copy_paste_flags = (short) IVisio.VisCutCopyPasteCodes.visCopyPasteNoTranslate;

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
            int num_src_shapes = src_page_shapes.Count;

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
            foreach (var src in _static_page_srcs)
            {
                query.Columns.Add(src, "Col" + i.ToString());
                i++;
            }

            var src_formulas = query.GetFormulas(src_page.PageSheet);

            // Set the Cells on the Destination

            var writer = new VASS.Writers.SrcWriter();
            for (i = 0; i < _static_page_srcs.Count; i++)
            {
                int row = 0;
                writer.SetValue(_static_page_srcs[i], src_formulas[row][i]);
            }

            writer.Commit(dest_page.PageSheet, Core.CellValueType.Formula);

            // make sure the new page looks like the old page
            dest_page.Background = src_page.Background;

            // then paste any contents from the first page
            if (num_src_shapes > 0)
            {
                dest_page.Paste(copy_paste_flags);
            }
        }

        private static void init_page_srcs()
        {
            if (_static_page_srcs == null)
            {
                _static_page_srcs = new List<Core.Src>();

                _static_page_srcs.Add(Core.SrcConstants.PrintLeftMargin);
                _static_page_srcs.Add(Core.SrcConstants.PrintCenterX);
                _static_page_srcs.Add(Core.SrcConstants.PrintCenterY);
                _static_page_srcs.Add(Core.SrcConstants.PrintOnPage);
                _static_page_srcs.Add(Core.SrcConstants.PrintBottomMargin);
                _static_page_srcs.Add(Core.SrcConstants.PrintRightMargin);
                _static_page_srcs.Add(Core.SrcConstants.PrintPagesX);
                _static_page_srcs.Add(Core.SrcConstants.PrintPagesY);
                _static_page_srcs.Add(Core.SrcConstants.PrintTopMargin);
                _static_page_srcs.Add(Core.SrcConstants.PrintPaperKind);
                _static_page_srcs.Add(Core.SrcConstants.PrintGrid);
                _static_page_srcs.Add(Core.SrcConstants.PrintPageOrientation);
                _static_page_srcs.Add(Core.SrcConstants.PrintScaleX);
                _static_page_srcs.Add(Core.SrcConstants.PrintScaleY);
                _static_page_srcs.Add(Core.SrcConstants.PrintPaperSource);

                _static_page_srcs.Add(Core.SrcConstants.PageDrawingScale);
                _static_page_srcs.Add(Core.SrcConstants.PageDrawingScaleType);
                _static_page_srcs.Add(Core.SrcConstants.PageDrawingSizeType);
                _static_page_srcs.Add(Core.SrcConstants.PageInhibitSnap);
                _static_page_srcs.Add(Core.SrcConstants.PageHeight);
                _static_page_srcs.Add(Core.SrcConstants.PageScale);
                _static_page_srcs.Add(Core.SrcConstants.PageWidth);
                _static_page_srcs.Add(Core.SrcConstants.PageShadowObliqueAngle);
                _static_page_srcs.Add(Core.SrcConstants.PageShadowOffsetX);
                _static_page_srcs.Add(Core.SrcConstants.PageShadowOffsetY);
                _static_page_srcs.Add(Core.SrcConstants.PageShadowScaleFactor);
                _static_page_srcs.Add(Core.SrcConstants.PageShadowType);
                _static_page_srcs.Add(Core.SrcConstants.PageUIVisibility);
                _static_page_srcs.Add(Core.SrcConstants.PageDrawingResizeType);

                _static_page_srcs.Add(Core.SrcConstants.XGridDensity);
                _static_page_srcs.Add(Core.SrcConstants.XGridOrigin);
                _static_page_srcs.Add(Core.SrcConstants.XGridSpacing);
                _static_page_srcs.Add(Core.SrcConstants.XRulerDensity);
                _static_page_srcs.Add(Core.SrcConstants.XRulerOrigin);
                _static_page_srcs.Add(Core.SrcConstants.YGridDensity);
                _static_page_srcs.Add(Core.SrcConstants.YGridOrigin);
                _static_page_srcs.Add(Core.SrcConstants.YGridSpacing);
                _static_page_srcs.Add(Core.SrcConstants.YRulerDensity);
                _static_page_srcs.Add(Core.SrcConstants.YRulerOrigin);

                _static_page_srcs.Add(Core.SrcConstants.PageLayoutAvenueSizeX);
                _static_page_srcs.Add(Core.SrcConstants.PageLayoutAvenueSizeY);
                _static_page_srcs.Add(Core.SrcConstants.PageLayoutBlockSizeX);
                _static_page_srcs.Add(Core.SrcConstants.PageLayoutBlockSizeY);
                _static_page_srcs.Add(Core.SrcConstants.PageLayoutControlAsInput);
                _static_page_srcs.Add(Core.SrcConstants.PageLayoutDynamicsOff);
                _static_page_srcs.Add(Core.SrcConstants.PageLayoutEnableGrid);
                _static_page_srcs.Add(Core.SrcConstants.PageLayoutLineAdjustFrom);
                _static_page_srcs.Add(Core.SrcConstants.PageLayoutLineAdjustTo);
                _static_page_srcs.Add(Core.SrcConstants.PageLayoutLineJumpCode);
                _static_page_srcs.Add(Core.SrcConstants.PageLayoutLineJumpFactorX);
                _static_page_srcs.Add(Core.SrcConstants.PageLayoutLineJumpFactorY);
                _static_page_srcs.Add(Core.SrcConstants.PageLayoutLineJumpStyle);
                _static_page_srcs.Add(Core.SrcConstants.PageLayoutLineRouteExt);
                _static_page_srcs.Add(Core.SrcConstants.PageLayoutLineToLineX);
                _static_page_srcs.Add(Core.SrcConstants.PageLayoutLineToLineY);
                _static_page_srcs.Add(Core.SrcConstants.PageLayoutLineToNodeX);
                _static_page_srcs.Add(Core.SrcConstants.PageLayoutLineToNodeY);
                _static_page_srcs.Add(Core.SrcConstants.PageLayoutLineJumpDirX);
                _static_page_srcs.Add(Core.SrcConstants.PageLayoutLineJumpDirY);
                _static_page_srcs.Add(Core.SrcConstants.PageLayoutShapeSplit);
                _static_page_srcs.Add(Core.SrcConstants.PageLayoutPlaceDepth);
                _static_page_srcs.Add(Core.SrcConstants.PageLayoutPlaceFlip);
                _static_page_srcs.Add(Core.SrcConstants.PageLayoutPlaceStyle);
                _static_page_srcs.Add(Core.SrcConstants.PageLayoutPlowCode);
                _static_page_srcs.Add(Core.SrcConstants.PageLayoutResizePage);
                _static_page_srcs.Add(Core.SrcConstants.PageLayoutRouteStyle);
                _static_page_srcs.Add(Core.SrcConstants.PageLayoutAvoidPageBreaks);
            }
        }

        public static Core.Size GetSize(IVisio.Page page)
        {
            var query = new VASS.Query.CellQuery();
            var col_height = query.Columns.Add(Core.SrcConstants.PageHeight);
            var col_width = query.Columns.Add(Core.SrcConstants.PageWidth);

            var cellqueryresult = query.GetResults<double>(page.PageSheet);
            var row = cellqueryresult[0];
            double height = row[col_height];
            double width = row[col_width];
            var s = new Core.Size(width, height);
            return s;
        }

        public static void SetSize(IVisio.Page page, Core.Size size)
        {
            var writer = new VASS.Writers.SrcWriter();
            writer.SetValue(Core.SrcConstants.PageWidth, size.Width);
            writer.SetValue(Core.SrcConstants.PageHeight, size.Height);

            writer.Commit(page.PageSheet, Core.CellValueType.Formula);
        }

        public static short[] DropManyAutoConnectors(
            IVisio.Page page,
            ICollection<Core.Point> points)
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
            var xy_array = Core.Point.ToDoubles(points).ToArray();

            System.Array outids_sa;

            page.DropManyU(masters_obj_array, xy_array, out outids_sa);

            short[] outids = (short[]) outids_sa;
            return outids;
        }

        public static void ResizeToFitContents(IVisio.Page page, Core.Size padding)
        {
            // first perform the native resizetofit
            page.ResizeToFitContents();

            if ((padding.Width > 0.0) || (padding.Height > 0.0))
            {
                // if there is any additional padding requested
                // we need to further handle the page

                // first determine the desired page size including the padding
                // and set the new size

                var old_size = GetSize(page);
                var new_size = old_size + padding.Multiply(2, 2);
                SetSize(page, new_size);

                // The page has the correct size, but
                // the contents will be offset from the correct location
                page.CenterDrawing();
            }
        }
    }
}