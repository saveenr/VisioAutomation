namespace VisioAutomation.Pages;

public static class PageHelper
{
    private static List<VASS.Src> _static_page_srcs;

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
        foreach (var src in _static_page_srcs)
        {
            query.Columns.Add(src,"Col"+i.ToString());
            i++;
        }

        var src_formulas = query.GetFormulas(src_page.PageSheet);

        // Set the Cells on the Destination
           
        var writer = new VASS.Writers.SrcWriter();
        for (i = 0; i < _static_page_srcs.Count; i++)
        {
            int row = 0;
            writer.SetValue(_static_page_srcs[i],src_formulas[row][i]);
        }

        writer.Commit(dest_page.PageSheet, VASS.CellValueType.Formula);

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
        if (_static_page_srcs == null)
        {
            _static_page_srcs = new List<VASS.Src>();

            _static_page_srcs.Add(VASS.SrcConstants.PrintLeftMargin);
            _static_page_srcs.Add(VASS.SrcConstants.PrintCenterX);
            _static_page_srcs.Add(VASS.SrcConstants.PrintCenterY);
            _static_page_srcs.Add(VASS.SrcConstants.PrintOnPage);
            _static_page_srcs.Add(VASS.SrcConstants.PrintBottomMargin);
            _static_page_srcs.Add(VASS.SrcConstants.PrintRightMargin);
            _static_page_srcs.Add(VASS.SrcConstants.PrintPagesX);
            _static_page_srcs.Add(VASS.SrcConstants.PrintPagesY);
            _static_page_srcs.Add(VASS.SrcConstants.PrintTopMargin);
            _static_page_srcs.Add(VASS.SrcConstants.PrintPaperKind);
            _static_page_srcs.Add(VASS.SrcConstants.PrintGrid);
            _static_page_srcs.Add(VASS.SrcConstants.PrintPageOrientation);
            _static_page_srcs.Add(VASS.SrcConstants.PrintScaleX);
            _static_page_srcs.Add(VASS.SrcConstants.PrintScaleY);
            _static_page_srcs.Add(VASS.SrcConstants.PrintPaperSource);

            _static_page_srcs.Add(VASS.SrcConstants.PageDrawingScale);
            _static_page_srcs.Add(VASS.SrcConstants.PageDrawingScaleType);
            _static_page_srcs.Add(VASS.SrcConstants.PageDrawingSizeType);
            _static_page_srcs.Add(VASS.SrcConstants.PageInhibitSnap);
            _static_page_srcs.Add(VASS.SrcConstants.PageHeight);
            _static_page_srcs.Add(VASS.SrcConstants.PageScale);
            _static_page_srcs.Add(VASS.SrcConstants.PageWidth);
            _static_page_srcs.Add(VASS.SrcConstants.PageShadowObliqueAngle);
            _static_page_srcs.Add(VASS.SrcConstants.PageShadowOffsetX);
            _static_page_srcs.Add(VASS.SrcConstants.PageShadowOffsetY);
            _static_page_srcs.Add(VASS.SrcConstants.PageShadowScaleFactor);
            _static_page_srcs.Add(VASS.SrcConstants.PageShadowType);
            _static_page_srcs.Add(VASS.SrcConstants.PageUIVisibility);
            _static_page_srcs.Add(VASS.SrcConstants.PageDrawingResizeType);

            _static_page_srcs.Add(VASS.SrcConstants.XGridDensity);
            _static_page_srcs.Add(VASS.SrcConstants.XGridOrigin);
            _static_page_srcs.Add(VASS.SrcConstants.XGridSpacing);
            _static_page_srcs.Add(VASS.SrcConstants.XRulerDensity);
            _static_page_srcs.Add(VASS.SrcConstants.XRulerOrigin);
            _static_page_srcs.Add(VASS.SrcConstants.YGridDensity);
            _static_page_srcs.Add(VASS.SrcConstants.YGridOrigin);
            _static_page_srcs.Add(VASS.SrcConstants.YGridSpacing);
            _static_page_srcs.Add(VASS.SrcConstants.YRulerDensity);
            _static_page_srcs.Add(VASS.SrcConstants.YRulerOrigin);

            _static_page_srcs.Add(VASS.SrcConstants.PageLayoutAvenueSizeX);
            _static_page_srcs.Add(VASS.SrcConstants.PageLayoutAvenueSizeY);
            _static_page_srcs.Add(VASS.SrcConstants.PageLayoutBlockSizeX);
            _static_page_srcs.Add(VASS.SrcConstants.PageLayoutBlockSizeY);
            _static_page_srcs.Add(VASS.SrcConstants.PageLayoutControlAsInput);
            _static_page_srcs.Add(VASS.SrcConstants.PageLayoutDynamicsOff);
            _static_page_srcs.Add(VASS.SrcConstants.PageLayoutEnableGrid);
            _static_page_srcs.Add(VASS.SrcConstants.PageLayoutLineAdjustFrom);
            _static_page_srcs.Add(VASS.SrcConstants.PageLayoutLineAdjustTo);
            _static_page_srcs.Add(VASS.SrcConstants.PageLayoutLineJumpCode);
            _static_page_srcs.Add(VASS.SrcConstants.PageLayoutLineJumpFactorX);
            _static_page_srcs.Add(VASS.SrcConstants.PageLayoutLineJumpFactorY);
            _static_page_srcs.Add(VASS.SrcConstants.PageLayoutLineJumpStyle);
            _static_page_srcs.Add(VASS.SrcConstants.PageLayoutLineRouteExt);
            _static_page_srcs.Add(VASS.SrcConstants.PageLayoutLineToLineX);
            _static_page_srcs.Add(VASS.SrcConstants.PageLayoutLineToLineY);
            _static_page_srcs.Add(VASS.SrcConstants.PageLayoutLineToNodeX);
            _static_page_srcs.Add(VASS.SrcConstants.PageLayoutLineToNodeY);
            _static_page_srcs.Add(VASS.SrcConstants.PageLayoutLineJumpDirX);
            _static_page_srcs.Add(VASS.SrcConstants.PageLayoutLineJumpDirY);
            _static_page_srcs.Add(VASS.SrcConstants.PageLayoutShapeSplit);
            _static_page_srcs.Add(VASS.SrcConstants.PageLayoutPlaceDepth);
            _static_page_srcs.Add(VASS.SrcConstants.PageLayoutPlaceFlip);
            _static_page_srcs.Add(VASS.SrcConstants.PageLayoutPlaceStyle);
            _static_page_srcs.Add(VASS.SrcConstants.PageLayoutPlowCode);
            _static_page_srcs.Add(VASS.SrcConstants.PageLayoutResizePage);
            _static_page_srcs.Add(VASS.SrcConstants.PageLayoutRouteStyle);
            _static_page_srcs.Add(VASS.SrcConstants.PageLayoutAvoidPageBreaks);
        }
    }

    public static Geometry.Size GetSize(IVisio.Page page)
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

    public static void SetSize(IVisio.Page page, Geometry.Size size)
    {
        var writer = new VASS.Writers.SrcWriter();
        writer.SetValue(VASS.SrcConstants.PageWidth, size.Width);
        writer.SetValue(VASS.SrcConstants.PageHeight, size.Height);

        writer.Commit(page.PageSheet, VASS.CellValueType.Formula);
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

    public static void ResizeToFitContents(IVisio.Page page, Geometry.Size padding)
    {
        // first perform the native resizetofit
        page.ResizeToFitContents();

        if ((padding.Width > 0.0) || (padding.Height > 0.0))
        {
            // if there is any additional padding requested
            // we need to further handle the page

            // first determine the desired page size including the padding
            // and set the new size

            var old_size = VisioAutomation.Pages.PageHelper.GetSize(page);
            var new_size = old_size + padding.Multiply(2, 2);
            VisioAutomation.Pages.PageHelper.SetSize(page, new_size);

            // The page has the correct size, but
            // the contents will be offset from the correct location
            page.CenterDrawing();
        }
    }

    public static short[] DropManyU(
        IVisio.Page page,
        IList<IVisio.Master> masters,
        IEnumerable<Geometry.Point> points)
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
        var xy_array = Geometry.Point.ToDoubles(points).ToArray();

        System.Array outids_sa;

        page.DropManyU(masters_obj_array, xy_array, out outids_sa);

        short[] outids = (short[])outids_sa;
        return outids;
    }
}