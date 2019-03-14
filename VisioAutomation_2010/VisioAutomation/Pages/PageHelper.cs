using System.Collections.Generic;
using System.Linq;
using VASS=VisioAutomation.ShapeSheet;
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
            var query = new VASS.Query.CellQuery();
            int i = 0;
            foreach (var src in page_srcs)
            {
                query.Columns.Add(src,"Col"+i.ToString());
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
                page_srcs = new List<VASS.Src>();

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

        internal static Geometry.Size GetSize(IVisio.Page page)
        {
            var query = new VASS.Query.CellQuery();
            var col_height = query.Columns.Add(ShapeSheet.SrcConstants.PageHeight,nameof(ShapeSheet.SrcConstants.PageHeight));
            var col_width = query.Columns.Add(ShapeSheet.SrcConstants.PageWidth,nameof(ShapeSheet.SrcConstants.PageWidth));

            var results = query.GetResults<double>(page.PageSheet);
            double height = results.Cells[col_height];
            double width = results.Cells[col_width];
            var s = new Geometry.Size(width, height);
            return s;
        }

        internal static void SetSize(IVisio.Page page, Geometry.Size size)
        {
            var writer = new VisioAutomation.ShapeSheet.Writers.SrcWriter();
            writer.SetFormula(VisioAutomation.ShapeSheet.SrcConstants.PageWidth, size.Width);
            writer.SetFormula(VisioAutomation.ShapeSheet.SrcConstants.PageHeight, size.Height);

            writer.Commit(page.PageSheet);
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


        public static PageRulerAndGridCells GetPageRulerAndGridCells(IVisio.Shape shape, VASS.CellValueType type)
        {
            var reader = PageRulerAndGridCells_lazy_reader.Value;
            return reader.GetCellsSingleRow(shape, type);
        }

        private static readonly System.Lazy<PageRulerAndGridCellsReader> PageRulerAndGridCells_lazy_reader = new System.Lazy<PageRulerAndGridCellsReader>();

        class PageRulerAndGridCellsReader : VASS.CellGroups.CellGroupReader<PageRulerAndGridCells>
        {
            public VASS.Query.CellColumn XGridDensity { get; set; }
            public VASS.Query.CellColumn XGridOrigin { get; set; }
            public VASS.Query.CellColumn XGridSpacing { get; set; }
            public VASS.Query.CellColumn XRulerDensity { get; set; }
            public VASS.Query.CellColumn XRulerOrigin { get; set; }
            public VASS.Query.CellColumn YGridDensity { get; set; }
            public VASS.Query.CellColumn YGridOrigin { get; set; }
            public VASS.Query.CellColumn YGridSpacing { get; set; }
            public VASS.Query.CellColumn YRulerDensity { get; set; }
            public VASS.Query.CellColumn YRulerOrigin { get; set; }

            public PageRulerAndGridCellsReader() : base(new VisioAutomation.ShapeSheet.Query.CellQuery())
            {
                this.XGridDensity = this.query_singlerow.Columns.Add(VASS.SrcConstants.XGridDensity, nameof(this.XGridDensity));
                this.XGridOrigin = this.query_singlerow.Columns.Add(VASS.SrcConstants.XGridOrigin, nameof(this.XGridOrigin));
                this.XGridSpacing = this.query_singlerow.Columns.Add(VASS.SrcConstants.XGridSpacing, nameof(this.XGridSpacing));
                this.XRulerDensity = this.query_singlerow.Columns.Add(VASS.SrcConstants.XRulerDensity, nameof(this.XRulerDensity));
                this.XRulerOrigin = this.query_singlerow.Columns.Add(VASS.SrcConstants.XRulerOrigin, nameof(this.XRulerOrigin));
                this.YGridDensity = this.query_singlerow.Columns.Add(VASS.SrcConstants.YGridDensity, nameof(this.YGridDensity));
                this.YGridOrigin = this.query_singlerow.Columns.Add(VASS.SrcConstants.YGridOrigin, nameof(this.YGridOrigin));
                this.YGridSpacing = this.query_singlerow.Columns.Add(VASS.SrcConstants.YGridSpacing, nameof(this.YGridSpacing));
                this.YRulerDensity = this.query_singlerow.Columns.Add(VASS.SrcConstants.YRulerDensity, nameof(this.YRulerDensity));
                this.YRulerOrigin = this.query_singlerow.Columns.Add(VASS.SrcConstants.YRulerOrigin, nameof(this.YRulerOrigin));
            }

            public override PageRulerAndGridCells ToCellGroup(ShapeSheet.Internal.ArraySegment<string> row)
            {
                var cells = new PageRulerAndGridCells();
                cells.XGridDensity = row[this.XGridDensity];
                cells.XGridOrigin = row[this.XGridOrigin];
                cells.XGridSpacing = row[this.XGridSpacing];
                cells.XRulerDensity = row[this.XRulerDensity];
                cells.XRulerOrigin = row[this.XRulerOrigin];
                cells.YGridDensity = row[this.YGridDensity];
                cells.YGridOrigin = row[this.YGridOrigin];
                cells.YGridSpacing = row[this.YGridSpacing];
                cells.YRulerDensity = row[this.YRulerDensity];
                cells.YRulerOrigin = row[this.YRulerOrigin];
                return cells;
            }
        }

        public static PageFormatCells GetPageFormatCells(IVisio.Shape shape, VASS.CellValueType type)
        {
            var reader = PageFormatCells_lazy_reader.Value;
            return reader.GetCellsSingleRow(shape, type);
        }

        private static readonly System.Lazy<PageFormatCellsReader> PageFormatCells_lazy_reader = new System.Lazy<PageFormatCellsReader>();

        class PageFormatCellsReader : VASS.CellGroups.CellGroupReader<PageFormatCells>
        {
            public VASS.Query.CellColumn DrawingScale { get; set; }
            public VASS.Query.CellColumn DrawingScaleType { get; set; }
            public VASS.Query.CellColumn DrawingSizeType { get; set; }
            public VASS.Query.CellColumn InhibitSnap { get; set; }
            public VASS.Query.CellColumn Height { get; set; }
            public VASS.Query.CellColumn Scale { get; set; }
            public VASS.Query.CellColumn Width { get; set; }
            public VASS.Query.CellColumn ShadowObliqueAngle { get; set; }
            public VASS.Query.CellColumn ShadowOffsetX { get; set; }
            public VASS.Query.CellColumn ShadowOffsetY { get; set; }
            public VASS.Query.CellColumn ShadowScaleFactor { get; set; }
            public VASS.Query.CellColumn ShadowType { get; set; }
            public VASS.Query.CellColumn UIVisibility { get; set; }
            public VASS.Query.CellColumn DrawingResizeType { get; set; }

            public PageFormatCellsReader() : base(new VisioAutomation.ShapeSheet.Query.CellQuery())
            {
                this.DrawingScale = this.query_singlerow.Columns.Add(VASS.SrcConstants.PageDrawingScale, nameof(this.DrawingScale));
                this.DrawingScaleType = this.query_singlerow.Columns.Add(VASS.SrcConstants.PageDrawingScaleType, nameof(this.DrawingScaleType));
                this.DrawingSizeType = this.query_singlerow.Columns.Add(VASS.SrcConstants.PageDrawingSizeType, nameof(this.DrawingSizeType));
                this.InhibitSnap = this.query_singlerow.Columns.Add(VASS.SrcConstants.PageInhibitSnap, nameof(this.InhibitSnap));
                this.Height = this.query_singlerow.Columns.Add(VASS.SrcConstants.PageHeight, nameof(this.Height));
                this.Scale = this.query_singlerow.Columns.Add(VASS.SrcConstants.PageScale, nameof(this.Scale));
                this.Width = this.query_singlerow.Columns.Add(VASS.SrcConstants.PageWidth, nameof(this.Width));
                this.ShadowObliqueAngle = this.query_singlerow.Columns.Add(VASS.SrcConstants.PageShadowObliqueAngle, nameof(this.ShadowObliqueAngle));
                this.ShadowOffsetX = this.query_singlerow.Columns.Add(VASS.SrcConstants.PageShadowOffsetX, nameof(this.ShadowOffsetX));
                this.ShadowOffsetY = this.query_singlerow.Columns.Add(VASS.SrcConstants.PageShadowOffsetY, nameof(this.ShadowOffsetY));
                this.ShadowScaleFactor = this.query_singlerow.Columns.Add(VASS.SrcConstants.PageShadowScaleFactor, nameof(this.ShadowScaleFactor));
                this.ShadowType = this.query_singlerow.Columns.Add(VASS.SrcConstants.PageShadowType, nameof(this.ShadowType));
                this.UIVisibility = this.query_singlerow.Columns.Add(VASS.SrcConstants.PageUIVisibility, nameof(this.UIVisibility));
                this.DrawingResizeType = this.query_singlerow.Columns.Add(VASS.SrcConstants.PageDrawingResizeType, nameof(this.DrawingResizeType));
            }

            public override PageFormatCells ToCellGroup(ShapeSheet.Internal.ArraySegment<string> row)
            {
                var cells = new PageFormatCells();
                cells.DrawingScale = row[this.DrawingScale];
                cells.DrawingScaleType = row[this.DrawingScaleType];
                cells.DrawingSizeType = row[this.DrawingSizeType];
                cells.InhibitSnap = row[this.InhibitSnap];
                cells.Height = row[this.Height];
                cells.Scale = row[this.Scale];
                cells.Width = row[this.Width];
                cells.ShadowObliqueAngle = row[this.ShadowObliqueAngle];
                cells.ShadowOffsetX = row[this.ShadowOffsetX];
                cells.ShadowOffsetY = row[this.ShadowOffsetY];
                cells.ShadowScaleFactor = row[this.ShadowScaleFactor];
                cells.ShadowType = row[this.ShadowType];
                cells.UIVisibility = row[this.UIVisibility];
                cells.DrawingResizeType = row[this.DrawingResizeType];
                return cells;
            }
        }

        public static PageLayoutCells GetPageLayoutCells(IVisio.Shape shape, VASS.CellValueType type)
        {
            var reader = PageLayoutCells_lazy_reader.Value;
            return reader.GetCellsSingleRow(shape, type);
        }

        private static readonly System.Lazy<PageLayoutCellsReader> PageLayoutCells_lazy_reader = new System.Lazy<PageLayoutCellsReader>();

        class PageLayoutCellsReader : VASS.CellGroups.CellGroupReader<PageLayoutCells>
        {
            public VASS.Query.CellColumn AvenueSizeX { get; set; }
            public VASS.Query.CellColumn AvenueSizeY { get; set; }
            public VASS.Query.CellColumn BlockSizeX { get; set; }
            public VASS.Query.CellColumn BlockSizeY { get; set; }
            public VASS.Query.CellColumn ControlAsInput { get; set; }
            public VASS.Query.CellColumn DynamicsOff { get; set; }
            public VASS.Query.CellColumn EnableGrid { get; set; }
            public VASS.Query.CellColumn LineAdjustFrom { get; set; }
            public VASS.Query.CellColumn LineAdjustTo { get; set; }
            public VASS.Query.CellColumn LineJumpCode { get; set; }
            public VASS.Query.CellColumn LineJumpFactorX { get; set; }
            public VASS.Query.CellColumn LineJumpFactorY { get; set; }
            public VASS.Query.CellColumn LineJumpStyle { get; set; }
            public VASS.Query.CellColumn LineRouteExt { get; set; }
            public VASS.Query.CellColumn LineToLineX { get; set; }
            public VASS.Query.CellColumn LineToLineY { get; set; }
            public VASS.Query.CellColumn LineToNodeX { get; set; }
            public VASS.Query.CellColumn LineToNodeY { get; set; }
            public VASS.Query.CellColumn LineJumpDirX { get; set; }
            public VASS.Query.CellColumn LineJumpDirY { get; set; }
            public VASS.Query.CellColumn ShapeSplit { get; set; }
            public VASS.Query.CellColumn PlaceDepth { get; set; }
            public VASS.Query.CellColumn PlaceFlip { get; set; }
            public VASS.Query.CellColumn PlaceStyle { get; set; }
            public VASS.Query.CellColumn PlowCode { get; set; }
            public VASS.Query.CellColumn ResizePage { get; set; }
            public VASS.Query.CellColumn RouteStyle { get; set; }
            public VASS.Query.CellColumn AvoidPageBreaks { get; set; }

            public PageLayoutCellsReader() : base(new VisioAutomation.ShapeSheet.Query.CellQuery())
            {
                this.AvenueSizeX = this.query_singlerow.Columns.Add(VASS.SrcConstants.PageLayoutAvenueSizeX, nameof(this.AvenueSizeX));
                this.AvenueSizeY = this.query_singlerow.Columns.Add(VASS.SrcConstants.PageLayoutAvenueSizeY, nameof(this.AvenueSizeY));
                this.BlockSizeX = this.query_singlerow.Columns.Add(VASS.SrcConstants.PageLayoutBlockSizeX, nameof(this.BlockSizeX));
                this.BlockSizeY = this.query_singlerow.Columns.Add(VASS.SrcConstants.PageLayoutBlockSizeY, nameof(this.BlockSizeY));
                this.ControlAsInput = this.query_singlerow.Columns.Add(VASS.SrcConstants.PageLayoutControlAsInput, nameof(this.ControlAsInput));
                this.DynamicsOff = this.query_singlerow.Columns.Add(VASS.SrcConstants.PageLayoutDynamicsOff, nameof(this.DynamicsOff));
                this.EnableGrid = this.query_singlerow.Columns.Add(VASS.SrcConstants.PageLayoutEnableGrid, nameof(this.EnableGrid));
                this.LineAdjustFrom = this.query_singlerow.Columns.Add(VASS.SrcConstants.PageLayoutLineAdjustFrom, nameof(this.LineAdjustFrom));
                this.LineAdjustTo = this.query_singlerow.Columns.Add(VASS.SrcConstants.PageLayoutLineAdjustTo, nameof(this.LineAdjustTo));
                this.LineJumpCode = this.query_singlerow.Columns.Add(VASS.SrcConstants.PageLayoutLineJumpCode, nameof(this.LineJumpCode));
                this.LineJumpFactorX = this.query_singlerow.Columns.Add(VASS.SrcConstants.PageLayoutLineJumpFactorX, nameof(this.LineJumpFactorX));
                this.LineJumpFactorY = this.query_singlerow.Columns.Add(VASS.SrcConstants.PageLayoutLineJumpFactorY, nameof(this.LineJumpFactorY));
                this.LineJumpStyle = this.query_singlerow.Columns.Add(VASS.SrcConstants.PageLayoutLineJumpStyle, nameof(this.LineJumpStyle));
                this.LineRouteExt = this.query_singlerow.Columns.Add(VASS.SrcConstants.PageLayoutLineRouteExt, nameof(this.LineRouteExt));
                this.LineToLineX = this.query_singlerow.Columns.Add(VASS.SrcConstants.PageLayoutLineToLineX, nameof(this.LineToLineX));
                this.LineToLineY = this.query_singlerow.Columns.Add(VASS.SrcConstants.PageLayoutLineToLineY, nameof(this.LineToLineY));
                this.LineToNodeX = this.query_singlerow.Columns.Add(VASS.SrcConstants.PageLayoutLineToNodeX, nameof(this.LineToNodeX));
                this.LineToNodeY = this.query_singlerow.Columns.Add(VASS.SrcConstants.PageLayoutLineToNodeY, nameof(this.LineToNodeY));
                this.LineJumpDirX = this.query_singlerow.Columns.Add(VASS.SrcConstants.PageLayoutLineJumpDirX, nameof(this.LineJumpDirX));
                this.LineJumpDirY = this.query_singlerow.Columns.Add(VASS.SrcConstants.PageLayoutLineJumpDirY, nameof(this.LineJumpDirY));
                this.ShapeSplit = this.query_singlerow.Columns.Add(VASS.SrcConstants.PageLayoutShapeSplit, nameof(this.ShapeSplit));
                this.PlaceDepth = this.query_singlerow.Columns.Add(VASS.SrcConstants.PageLayoutPlaceDepth, nameof(this.PlaceDepth));
                this.PlaceFlip = this.query_singlerow.Columns.Add(VASS.SrcConstants.PageLayoutPlaceFlip, nameof(this.PlaceFlip));
                this.PlaceStyle = this.query_singlerow.Columns.Add(VASS.SrcConstants.PageLayoutPlaceStyle, nameof(this.PlaceStyle));
                this.PlowCode = this.query_singlerow.Columns.Add(VASS.SrcConstants.PageLayoutPlowCode, nameof(this.PlowCode));
                this.ResizePage = this.query_singlerow.Columns.Add(VASS.SrcConstants.PageLayoutResizePage, nameof(this.ResizePage));
                this.RouteStyle = this.query_singlerow.Columns.Add(VASS.SrcConstants.PageLayoutRouteStyle, nameof(this.RouteStyle));
                this.AvoidPageBreaks = this.query_singlerow.Columns.Add(VASS.SrcConstants.PageLayoutAvoidPageBreaks, nameof(this.AvoidPageBreaks));
            }


            public override PageLayoutCells ToCellGroup(ShapeSheet.Internal.ArraySegment<string> row)
            {
                var cells = new PageLayoutCells();
                cells.AvenueSizeX = row[this.AvenueSizeX];
                cells.AvenueSizeY = row[this.AvenueSizeY];
                cells.BlockSizeX = row[this.BlockSizeX];
                cells.BlockSizeY = row[this.BlockSizeY];
                cells.CtrlAsInput = row[this.ControlAsInput];
                cells.DynamicsOff = row[this.DynamicsOff];
                cells.EnableGrid = row[this.EnableGrid];
                cells.LineAdjustFrom = row[this.LineAdjustFrom];
                cells.LineAdjustTo = row[this.LineAdjustTo];
                cells.LineJumpCode = row[this.LineJumpCode];
                cells.LineJumpFactorX = row[this.LineJumpFactorX];
                cells.LineJumpFactorY = row[this.LineJumpFactorY];
                cells.LineJumpStyle = row[this.LineJumpStyle];
                cells.LineRouteExt = row[this.LineRouteExt];
                cells.LineToLineX = row[this.LineToLineX];
                cells.LineToLineY = row[this.LineToLineY];
                cells.LineToNodeX = row[this.LineToNodeX];
                cells.LineToNodeY = row[this.LineToNodeY];
                cells.LineJumpDirX = row[this.LineJumpDirX];
                cells.LineJumpDirY = row[this.LineJumpDirY];
                cells.PageShapeSplit = row[this.ShapeSplit];
                cells.PlaceDepth = row[this.PlaceDepth];
                cells.PlaceFlip = row[this.PlaceFlip];
                cells.PlaceStyle = row[this.PlaceStyle];
                cells.PlowCode = row[this.PlowCode];
                cells.ResizePage = row[this.ResizePage];
                cells.RouteStyle = row[this.RouteStyle];
                cells.AvoidPageBreaks = row[this.AvoidPageBreaks];
                return cells;
            }
        }

        public static PagePrintCells GetPagePrintCells(IVisio.Shape shape, VASS.CellValueType type)
        {
            var reader = PagePrintCells_lazy_reader.Value;
            return reader.GetCellsSingleRow(shape, type);
        }

        private static readonly System.Lazy<PagePrintCellsReader> PagePrintCells_lazy_reader = new System.Lazy<PagePrintCellsReader>();

        class PagePrintCellsReader : VASS.CellGroups.CellGroupReader<PagePrintCells>
        {
            public VASS.Query.CellColumn LeftMargin { get; set; }
            public VASS.Query.CellColumn CenterX { get; set; }
            public VASS.Query.CellColumn CenterY { get; set; }
            public VASS.Query.CellColumn OnPage { get; set; }
            public VASS.Query.CellColumn BottomMargin { get; set; }
            public VASS.Query.CellColumn RightMargin { get; set; }
            public VASS.Query.CellColumn PagesX { get; set; }
            public VASS.Query.CellColumn PagesY { get; set; }
            public VASS.Query.CellColumn TopMargin { get; set; }
            public VASS.Query.CellColumn PaperKind { get; set; }
            public VASS.Query.CellColumn Grid { get; set; }
            public VASS.Query.CellColumn PageOrientation { get; set; }
            public VASS.Query.CellColumn ScaleX { get; set; }
            public VASS.Query.CellColumn ScaleY { get; set; }
            public VASS.Query.CellColumn PaperSource { get; set; }

            public PagePrintCellsReader() : base(new VisioAutomation.ShapeSheet.Query.CellQuery())
            {
                this.LeftMargin = this.query_singlerow.Columns.Add(VASS.SrcConstants.PrintLeftMargin, nameof(this.LeftMargin));
                this.CenterX = this.query_singlerow.Columns.Add(VASS.SrcConstants.PrintCenterX, nameof(this.CenterX));
                this.CenterY = this.query_singlerow.Columns.Add(VASS.SrcConstants.PrintCenterY, nameof(this.CenterY));
                this.OnPage = this.query_singlerow.Columns.Add(VASS.SrcConstants.PrintOnPage, nameof(this.OnPage));
                this.BottomMargin = this.query_singlerow.Columns.Add(VASS.SrcConstants.PrintBottomMargin, nameof(this.BottomMargin));
                this.RightMargin = this.query_singlerow.Columns.Add(VASS.SrcConstants.PrintRightMargin, nameof(this.RightMargin));
                this.PagesX = this.query_singlerow.Columns.Add(VASS.SrcConstants.PrintPagesX, nameof(this.PagesX));
                this.PagesY = this.query_singlerow.Columns.Add(VASS.SrcConstants.PrintPagesY, nameof(this.PagesY));
                this.TopMargin = this.query_singlerow.Columns.Add(VASS.SrcConstants.PrintTopMargin, nameof(this.TopMargin));
                this.PaperKind = this.query_singlerow.Columns.Add(VASS.SrcConstants.PrintPaperKind, nameof(this.PaperKind));
                this.Grid = this.query_singlerow.Columns.Add(VASS.SrcConstants.PrintGrid, nameof(this.Grid));
                this.PageOrientation = this.query_singlerow.Columns.Add(VASS.SrcConstants.PrintPageOrientation, nameof(this.PageOrientation));
                this.ScaleX = this.query_singlerow.Columns.Add(VASS.SrcConstants.PrintScaleX, nameof(this.ScaleX));
                this.ScaleY = this.query_singlerow.Columns.Add(VASS.SrcConstants.PrintScaleY, nameof(this.ScaleY));
                this.PaperSource = this.query_singlerow.Columns.Add(VASS.SrcConstants.PrintPaperSource, nameof(this.PaperSource));
            }

            public override PagePrintCells ToCellGroup(ShapeSheet.Internal.ArraySegment<string> row)
            {
                var cells = new PagePrintCells();
                cells.LeftMargin = row[this.LeftMargin];
                cells.CenterX = row[this.CenterX];
                cells.CenterY = row[this.CenterY];
                cells.OnPage = row[this.OnPage];
                cells.BottomMargin = row[this.BottomMargin];
                cells.RightMargin = row[this.RightMargin];
                cells.PagesX = row[this.PagesX];
                cells.PagesY = row[this.PagesY];
                cells.TopMargin = row[this.TopMargin];
                cells.PaperKind = row[this.PaperKind];
                cells.Grid = row[this.Grid];
                cells.Orientation = row[this.PageOrientation];
                cells.ScaleX = row[this.ScaleX];
                cells.ScaleY = row[this.ScaleY];
                cells.PaperSource = row[this.PaperSource];
                return cells;
            }
        }

    }
}