using System.Collections.Generic;
using IVisio=Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;
using VisioAutomation.Extensions;

namespace VisioAutomation.Layout
{
    public partial class PageCells : VA.ShapeSheet.CellGroups.CellGroup
    {
        public VA.ShapeSheet.CellData<double> PageLeftMargin { get; set; }
        public VA.ShapeSheet.CellData<double> CenterX { get; set; }
        public VA.ShapeSheet.CellData<double> CenterY { get; set; }
        public VA.ShapeSheet.CellData<int> OnPage { get; set; }
        public VA.ShapeSheet.CellData<double> PageBottomMargin { get; set; }
        public VA.ShapeSheet.CellData<double> PageRightMargin { get; set; }
        public VA.ShapeSheet.CellData<double> PagesX { get; set; }
        public VA.ShapeSheet.CellData<double> PagesY { get; set; }
        public VA.ShapeSheet.CellData<double> PageTopMargin { get; set; }
        public VA.ShapeSheet.CellData<int> PaperKind { get; set; }
        public VA.ShapeSheet.CellData<int> PrintGrid { get; set; }
        public VA.ShapeSheet.CellData<int> PrintPageOrientation { get; set; }
        public VA.ShapeSheet.CellData<double> ScaleX { get; set; }
        public VA.ShapeSheet.CellData<double> ScaleY { get; set; }
        public VA.ShapeSheet.CellData<int> PaperSource { get; set; }
        public VA.ShapeSheet.CellData<double> DrawingScale { get; set; }
        public VA.ShapeSheet.CellData<int> DrawingScaleType { get; set; }
        public VA.ShapeSheet.CellData<int> DrawingSizeType { get; set; }
        public VA.ShapeSheet.CellData<int> InhibitSnap { get; set; }
        public VA.ShapeSheet.CellData<double> PageHeight { get; set; }
        public VA.ShapeSheet.CellData<double> PageScale { get; set; }
        public VA.ShapeSheet.CellData<double> PageWidth { get; set; }
        public VA.ShapeSheet.CellData<double> ShdwObliqueAngle { get; set; }
        public VA.ShapeSheet.CellData<double> ShdwOffsetX { get; set; }
        public VA.ShapeSheet.CellData<double> ShdwOffsetY { get; set; }
        public VA.ShapeSheet.CellData<double> ShdwScaleFactor { get; set; }
        public VA.ShapeSheet.CellData<int> ShdwType { get; set; }
        public VA.ShapeSheet.CellData<double> UIVisibility { get; set; }
        public VA.ShapeSheet.CellData<double> XGridDensity { get; set; }
        public VA.ShapeSheet.CellData<double> XGridOrigin { get; set; }
        public VA.ShapeSheet.CellData<double> XGridSpacing { get; set; }
        public VA.ShapeSheet.CellData<double> XRulerDensity { get; set; }
        public VA.ShapeSheet.CellData<double> XRulerOrigin { get; set; }
        public VA.ShapeSheet.CellData<double> YGridDensity { get; set; }
        public VA.ShapeSheet.CellData<double> YGridOrigin { get; set; }
        public VA.ShapeSheet.CellData<double> YGridSpacing { get; set; }
        public VA.ShapeSheet.CellData<double> YRulerDensity { get; set; }
        public VA.ShapeSheet.CellData<double> YRulerOrigin { get; set; }
        public VA.ShapeSheet.CellData<double> AvenueSizeX { get; set; }
        public VA.ShapeSheet.CellData<double> AvenueSizeY { get; set; }
        public VA.ShapeSheet.CellData<double> BlockSizeX { get; set; }
        public VA.ShapeSheet.CellData<double> BlockSizeY { get; set; }
        public VA.ShapeSheet.CellData<int> CtrlAsInput { get; set; }
        public VA.ShapeSheet.CellData<int> DynamicsOff { get; set; }
        public VA.ShapeSheet.CellData<int> EnableGrid { get; set; }
        public VA.ShapeSheet.CellData<int> LineAdjustFrom { get; set; }
        public VA.ShapeSheet.CellData<double> LineAdjustTo { get; set; }
        public VA.ShapeSheet.CellData<double> LineJumpCode { get; set; }
        public VA.ShapeSheet.CellData<double> LineJumpFactorX { get; set; }
        public VA.ShapeSheet.CellData<double> LineJumpFactorY { get; set; }
        public VA.ShapeSheet.CellData<int> LineJumpStyle { get; set; }
        public VA.ShapeSheet.CellData<double> LineRouteExt { get; set; }
        public VA.ShapeSheet.CellData<double> LineToLineX { get; set; }
        public VA.ShapeSheet.CellData<double> LineToLineY { get; set; }
        public VA.ShapeSheet.CellData<double> LineToNodeX { get; set; }
        public VA.ShapeSheet.CellData<double> LineToNodeY { get; set; }
        public VA.ShapeSheet.CellData<double> PageLineJumpDirX { get; set; }
        public VA.ShapeSheet.CellData<double> PageLineJumpDirY { get; set; }
        public VA.ShapeSheet.CellData<int> PageShapeSplit { get; set; }
        public VA.ShapeSheet.CellData<int> PlaceDepth { get; set; }
        public VA.ShapeSheet.CellData<int> PlaceFlip { get; set; }
        public VA.ShapeSheet.CellData<int> PlaceStyle { get; set; }
        public VA.ShapeSheet.CellData<int> PlowCode { get; set; }
        public VA.ShapeSheet.CellData<int> ResizePage { get; set; }
        public VA.ShapeSheet.CellData<int> RouteStyle { get; set; }

        protected override void _Apply(VA.ShapeSheet.CellGroups.CellGroup.ApplyFormula func)
        {
            func(ShapeSheet.SRCConstants.PageLeftMargin, this.PageLeftMargin.Formula);
            func(ShapeSheet.SRCConstants.CenterX, this.CenterX.Formula);
            func(ShapeSheet.SRCConstants.CenterY, this.CenterY.Formula);
            func(ShapeSheet.SRCConstants.OnPage, this.OnPage.Formula);
            func(ShapeSheet.SRCConstants.PageBottomMargin, this.PageBottomMargin.Formula);
            func(ShapeSheet.SRCConstants.PageRightMargin, this.PageRightMargin.Formula);
            func(ShapeSheet.SRCConstants.PagesX, this.PagesX.Formula);
            func(ShapeSheet.SRCConstants.PagesY, this.PagesY.Formula);
            func(ShapeSheet.SRCConstants.PageTopMargin, this.PageTopMargin.Formula);
            func(ShapeSheet.SRCConstants.PaperKind, this.PaperKind.Formula);
            func(ShapeSheet.SRCConstants.PrintGrid, this.PrintGrid.Formula);
            func(ShapeSheet.SRCConstants.PrintPageOrientation, this.PrintPageOrientation.Formula);
            func(ShapeSheet.SRCConstants.ScaleX, this.ScaleX.Formula);
            func(ShapeSheet.SRCConstants.ScaleY, this.ScaleY.Formula);
            func(ShapeSheet.SRCConstants.PaperSource, this.PaperSource.Formula);
            func(ShapeSheet.SRCConstants.DrawingScale, this.DrawingScale.Formula);
            func(ShapeSheet.SRCConstants.DrawingScaleType, this.DrawingScaleType.Formula);
            func(ShapeSheet.SRCConstants.DrawingSizeType, this.DrawingSizeType.Formula);
            func(ShapeSheet.SRCConstants.InhibitSnap, this.InhibitSnap.Formula);
            func(ShapeSheet.SRCConstants.PageHeight, this.PageHeight.Formula);
            func(ShapeSheet.SRCConstants.PageScale, this.PageScale.Formula);
            func(ShapeSheet.SRCConstants.PageWidth, this.PageWidth.Formula);
            func(ShapeSheet.SRCConstants.ShdwObliqueAngle, this.ShdwObliqueAngle.Formula);
            func(ShapeSheet.SRCConstants.ShdwOffsetX, this.ShdwOffsetX.Formula);
            func(ShapeSheet.SRCConstants.ShdwOffsetY, this.ShdwOffsetY.Formula);
            func(ShapeSheet.SRCConstants.ShdwScaleFactor, this.ShdwScaleFactor.Formula);
            func(ShapeSheet.SRCConstants.ShdwType, this.ShdwType.Formula);
            func(ShapeSheet.SRCConstants.UIVisibility, this.UIVisibility.Formula);
            func(ShapeSheet.SRCConstants.XGridDensity, this.XGridDensity.Formula);
            func(ShapeSheet.SRCConstants.XGridOrigin, this.XGridOrigin.Formula);
            func(ShapeSheet.SRCConstants.XGridSpacing, this.XGridSpacing.Formula);
            func(ShapeSheet.SRCConstants.XRulerDensity, this.XRulerDensity.Formula);
            func(ShapeSheet.SRCConstants.XRulerOrigin, this.XRulerOrigin.Formula);
            func(ShapeSheet.SRCConstants.YGridDensity, this.YGridDensity.Formula);
            func(ShapeSheet.SRCConstants.YGridOrigin, this.YGridOrigin.Formula);
            func(ShapeSheet.SRCConstants.YGridSpacing, this.YGridSpacing.Formula);
            func(ShapeSheet.SRCConstants.YRulerDensity, this.YRulerDensity.Formula);
            func(ShapeSheet.SRCConstants.YRulerOrigin, this.YRulerOrigin.Formula);
            func(ShapeSheet.SRCConstants.AvenueSizeX, this.AvenueSizeX.Formula);
            func(ShapeSheet.SRCConstants.AvenueSizeY, this.AvenueSizeY.Formula);
            func(ShapeSheet.SRCConstants.BlockSizeX, this.BlockSizeX.Formula);
            func(ShapeSheet.SRCConstants.BlockSizeY, this.BlockSizeY.Formula);
            func(ShapeSheet.SRCConstants.CtrlAsInput, this.CtrlAsInput.Formula);
            func(ShapeSheet.SRCConstants.DynamicsOff, this.DynamicsOff.Formula);
            func(ShapeSheet.SRCConstants.EnableGrid, this.EnableGrid.Formula);
            func(ShapeSheet.SRCConstants.LineAdjustFrom, this.LineAdjustFrom.Formula);
            func(ShapeSheet.SRCConstants.LineAdjustTo, this.LineAdjustTo.Formula);
            func(ShapeSheet.SRCConstants.LineJumpCode, this.LineJumpCode.Formula);
            func(ShapeSheet.SRCConstants.LineJumpFactorX, this.LineJumpFactorX.Formula);
            func(ShapeSheet.SRCConstants.LineJumpFactorY, this.LineJumpFactorY.Formula);
            func(ShapeSheet.SRCConstants.LineJumpStyle, this.LineJumpStyle.Formula);
            func(ShapeSheet.SRCConstants.LineRouteExt, this.LineRouteExt.Formula);
            func(ShapeSheet.SRCConstants.LineToLineX, this.LineToLineX.Formula);
            func(ShapeSheet.SRCConstants.LineToLineY, this.LineToLineY.Formula);
            func(ShapeSheet.SRCConstants.LineToNodeX, this.LineToNodeX.Formula);
            func(ShapeSheet.SRCConstants.LineToNodeY, this.LineToNodeY.Formula);
            func(ShapeSheet.SRCConstants.PageLineJumpDirX, this.PageLineJumpDirX.Formula);
            func(ShapeSheet.SRCConstants.PageLineJumpDirY, this.PageLineJumpDirY.Formula);
            func(ShapeSheet.SRCConstants.PageShapeSplit, this.PageShapeSplit.Formula);
            func(ShapeSheet.SRCConstants.PlaceDepth, this.PlaceDepth.Formula);
            func(ShapeSheet.SRCConstants.PlaceFlip, this.PlaceFlip.Formula);
            func(ShapeSheet.SRCConstants.PlaceStyle, this.PlaceStyle.Formula);
            func(ShapeSheet.SRCConstants.PlowCode, this.PlowCode.Formula);
            func(ShapeSheet.SRCConstants.ResizePage, this.ResizePage.Formula);
            func(ShapeSheet.SRCConstants.RouteStyle, this.RouteStyle.Formula);
        }

        private static PageCells get_cells_from_row(PageQuery query, VA.ShapeSheet.Data.QueryDataRow<double> row)
        {

            var cells = new PageCells();
            cells.PageLeftMargin = row[query.PageLeftMargin];
            cells.CenterX = row[query.CenterX];
            cells.CenterY = row[query.CenterY];
            cells.OnPage = row[query.OnPage].ToInt();
            cells.PageBottomMargin = row[query.PageBottomMargin];
            cells.PageRightMargin = row[query.PageRightMargin];
            cells.PagesX = row[query.PagesX];
            cells.PagesY = row[query.PagesY];
            cells.PageTopMargin = row[query.PageTopMargin];
            cells.PaperKind = row[query.PaperKind].ToInt();
            cells.PrintGrid = row[query.PrintGrid].ToInt();
            cells.PrintPageOrientation = row[query.PrintPageOrientation].ToInt();
            cells.ScaleX = row[query.ScaleX];
            cells.ScaleY = row[query.ScaleY];
            cells.PaperSource = row[query.PaperSource].ToInt();
            cells.DrawingScale = row[query.DrawingScale];
            cells.DrawingScaleType = row[query.DrawingScaleType].ToInt();
            cells.DrawingSizeType = row[query.DrawingSizeType].ToInt();
            cells.InhibitSnap = row[query.InhibitSnap].ToInt();
            cells.PageHeight = row[query.PageHeight];
            cells.PageScale = row[query.PageScale];
            cells.PageWidth = row[query.PageWidth];
            cells.ShdwObliqueAngle = row[query.ShdwObliqueAngle];
            cells.ShdwOffsetX = row[query.ShdwOffsetX];
            cells.ShdwOffsetY = row[query.ShdwOffsetY];
            cells.ShdwScaleFactor = row[query.ShdwScaleFactor];
            cells.ShdwType = row[query.ShdwType].ToInt();
            cells.UIVisibility = row[query.UIVisibility];
            cells.XGridDensity = row[query.XGridDensity];
            cells.XGridOrigin = row[query.XGridOrigin];
            cells.XGridSpacing = row[query.XGridSpacing];
            cells.XRulerDensity = row[query.XRulerDensity];
            cells.XRulerOrigin = row[query.XRulerOrigin];
            cells.YGridDensity = row[query.YGridDensity];
            cells.YGridOrigin = row[query.YGridOrigin];
            cells.YGridSpacing = row[query.YGridSpacing];
            cells.YRulerDensity = row[query.YRulerDensity];
            cells.YRulerOrigin = row[query.YRulerOrigin];
            cells.AvenueSizeX = row[query.AvenueSizeX];
            cells.AvenueSizeY = row[query.AvenueSizeY];
            cells.BlockSizeX = row[query.BlockSizeX];
            cells.BlockSizeY = row[query.BlockSizeY];
            cells.CtrlAsInput = row[query.CtrlAsInput].ToInt();
            cells.DynamicsOff = row[query.DynamicsOff].ToInt();
            cells.EnableGrid = row[query.EnableGrid].ToInt();
            cells.LineAdjustFrom = row[query.LineAdjustFrom].ToInt();
            cells.LineAdjustTo = row[query.LineAdjustTo];
            cells.LineJumpCode = row[query.LineJumpCode];
            cells.LineJumpFactorX = row[query.LineJumpFactorX];
            cells.LineJumpFactorY = row[query.LineJumpFactorY];
            cells.LineJumpStyle = row[query.LineJumpStyle].ToInt();
            cells.LineRouteExt = row[query.LineRouteExt];
            cells.LineToLineX = row[query.LineToLineX];
            cells.LineToLineY = row[query.LineToLineY];
            cells.LineToNodeX = row[query.LineToNodeX];
            cells.LineToNodeY = row[query.LineToNodeY];
            cells.PageLineJumpDirX = row[query.PageLineJumpDirX];
            cells.PageLineJumpDirY = row[query.PageLineJumpDirY];
            cells.PageShapeSplit = row[query.PageShapeSplit].ToInt();
            cells.PlaceDepth = row[query.PlaceDepth].ToInt();
            cells.PlaceFlip = row[query.PlaceFlip].ToInt();
            cells.PlaceStyle = row[query.PlaceStyle].ToInt();
            cells.PlowCode = row[query.PlowCode].ToInt();
            cells.ResizePage = row[query.ResizePage].ToInt();
            cells.RouteStyle = row[query.RouteStyle].ToInt();
            return cells;
        }

        internal static IList<PageCells> GetCells(IVisio.Page page, IList<int> shapeids)
        {
            var query = new PageQuery();
            return VA.ShapeSheet.CellGroups.CellGroup._GetObjectsFromRows(page, shapeids, query, get_cells_from_row);
        }

        internal static PageCells GetCells(IVisio.Shape shape)
        {
            var query = new PageQuery();
            return VA.ShapeSheet.CellGroups.CellGroup._GetObjectFromSingleRow(shape, query, get_cells_from_row);
        }

        class PageQuery : VA.ShapeSheet.Query.CellQuery
        {
            public VA.ShapeSheet.Query.CellQueryColumn PageLeftMargin { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn CenterX { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn CenterY { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn OnPage { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn PageBottomMargin { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn PageRightMargin { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn PagesX { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn PagesY { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn PageTopMargin { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn PaperKind { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn PrintGrid { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn PrintPageOrientation { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn ScaleX { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn ScaleY { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn PaperSource { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn DrawingScale { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn DrawingScaleType { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn DrawingSizeType { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn InhibitSnap { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn PageHeight { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn PageScale { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn PageWidth { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn ShdwObliqueAngle { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn ShdwOffsetX { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn ShdwOffsetY { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn ShdwScaleFactor { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn ShdwType { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn UIVisibility { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn XGridDensity { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn XGridOrigin { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn XGridSpacing { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn XRulerDensity { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn XRulerOrigin { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn YGridDensity { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn YGridOrigin { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn YGridSpacing { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn YRulerDensity { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn YRulerOrigin { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn AvenueSizeX { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn AvenueSizeY { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn BlockSizeX { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn BlockSizeY { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn CtrlAsInput { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn DynamicsOff { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn EnableGrid { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn LineAdjustFrom { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn LineAdjustTo { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn LineJumpCode { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn LineJumpFactorX { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn LineJumpFactorY { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn LineJumpStyle { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn LineRouteExt { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn LineToLineX { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn LineToLineY { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn LineToNodeX { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn LineToNodeY { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn PageLineJumpDirX { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn PageLineJumpDirY { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn PageShapeSplit { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn PlaceDepth { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn PlaceFlip { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn PlaceStyle { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn PlowCode { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn ResizePage { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn RouteStyle { get; set; }

            public PageQuery() :
                base()
            {
                this.PageLeftMargin = this.AddColumn(VA.ShapeSheet.SRCConstants.PageLeftMargin, "PageLeftMargin");
                this.CenterX = this.AddColumn(VA.ShapeSheet.SRCConstants.CenterX, "CenterX");
                this.CenterY = this.AddColumn(VA.ShapeSheet.SRCConstants.CenterY, "CenterY");
                this.OnPage = this.AddColumn(VA.ShapeSheet.SRCConstants.OnPage, "OnPage");
                this.PageBottomMargin = this.AddColumn(VA.ShapeSheet.SRCConstants.PageBottomMargin, "PageBottomMargin");
                this.PageRightMargin = this.AddColumn(VA.ShapeSheet.SRCConstants.PageRightMargin, "PageRightMargin");
                this.PagesX = this.AddColumn(VA.ShapeSheet.SRCConstants.PagesX, "PagesX");
                this.PagesY = this.AddColumn(VA.ShapeSheet.SRCConstants.PagesY, "PagesY");
                this.PageTopMargin = this.AddColumn(VA.ShapeSheet.SRCConstants.PageTopMargin, "PageTopMargin");
                this.PaperKind = this.AddColumn(VA.ShapeSheet.SRCConstants.PaperKind, "PaperKind");
                this.PrintGrid = this.AddColumn(VA.ShapeSheet.SRCConstants.PrintGrid, "PrintGrid");
                this.PrintPageOrientation = this.AddColumn(VA.ShapeSheet.SRCConstants.PrintPageOrientation, "PrintPageOrientation");
                this.ScaleX = this.AddColumn(VA.ShapeSheet.SRCConstants.ScaleX, "ScaleX");
                this.ScaleY = this.AddColumn(VA.ShapeSheet.SRCConstants.ScaleY, "ScaleY");
                this.PaperSource = this.AddColumn(VA.ShapeSheet.SRCConstants.PaperSource, "PaperSource");
                this.DrawingScale = this.AddColumn(VA.ShapeSheet.SRCConstants.DrawingScale, "DrawingScale");
                this.DrawingScaleType = this.AddColumn(VA.ShapeSheet.SRCConstants.DrawingScaleType, "DrawingScaleType");
                this.DrawingSizeType = this.AddColumn(VA.ShapeSheet.SRCConstants.DrawingSizeType, "DrawingSizeType");
                this.InhibitSnap = this.AddColumn(VA.ShapeSheet.SRCConstants.InhibitSnap, "InhibitSnap");
                this.PageHeight = this.AddColumn(VA.ShapeSheet.SRCConstants.PageHeight, "PageHeight");
                this.PageScale = this.AddColumn(VA.ShapeSheet.SRCConstants.PageScale, "PageScale");
                this.PageWidth = this.AddColumn(VA.ShapeSheet.SRCConstants.PageWidth, "PageWidth");
                this.ShdwObliqueAngle = this.AddColumn(VA.ShapeSheet.SRCConstants.ShdwObliqueAngle, "ShdwObliqueAngle");
                this.ShdwOffsetX = this.AddColumn(VA.ShapeSheet.SRCConstants.ShdwOffsetX, "ShdwOffsetX");
                this.ShdwOffsetY = this.AddColumn(VA.ShapeSheet.SRCConstants.ShdwOffsetY, "ShdwOffsetY");
                this.ShdwScaleFactor = this.AddColumn(VA.ShapeSheet.SRCConstants.ShdwScaleFactor, "ShdwScaleFactor");
                this.ShdwType = this.AddColumn(VA.ShapeSheet.SRCConstants.ShdwType, "ShdwType");
                this.UIVisibility = this.AddColumn(VA.ShapeSheet.SRCConstants.UIVisibility, "UIVisibility");
                this.XGridDensity = this.AddColumn(VA.ShapeSheet.SRCConstants.XGridDensity, "XGridDensity");
                this.XGridOrigin = this.AddColumn(VA.ShapeSheet.SRCConstants.XGridOrigin, "XGridOrigin");
                this.XGridSpacing = this.AddColumn(VA.ShapeSheet.SRCConstants.XGridSpacing, "XGridSpacing");
                this.XRulerDensity = this.AddColumn(VA.ShapeSheet.SRCConstants.XRulerDensity, "XRulerDensity");
                this.XRulerOrigin = this.AddColumn(VA.ShapeSheet.SRCConstants.XRulerOrigin, "XRulerOrigin");
                this.YGridDensity = this.AddColumn(VA.ShapeSheet.SRCConstants.YGridDensity, "YGridDensity");
                this.YGridOrigin = this.AddColumn(VA.ShapeSheet.SRCConstants.YGridOrigin, "YGridOrigin");
                this.YGridSpacing = this.AddColumn(VA.ShapeSheet.SRCConstants.YGridSpacing, "YGridSpacing");
                this.YRulerDensity = this.AddColumn(VA.ShapeSheet.SRCConstants.YRulerDensity, "YRulerDensity");
                this.YRulerOrigin = this.AddColumn(VA.ShapeSheet.SRCConstants.YRulerOrigin, "YRulerOrigin");
                this.AvenueSizeX = this.AddColumn(VA.ShapeSheet.SRCConstants.AvenueSizeX, "AvenueSizeX");
                this.AvenueSizeY = this.AddColumn(VA.ShapeSheet.SRCConstants.AvenueSizeY, "AvenueSizeY");
                this.BlockSizeX = this.AddColumn(VA.ShapeSheet.SRCConstants.BlockSizeX, "BlockSizeX");
                this.BlockSizeY = this.AddColumn(VA.ShapeSheet.SRCConstants.BlockSizeY, "BlockSizeY");
                this.CtrlAsInput = this.AddColumn(VA.ShapeSheet.SRCConstants.CtrlAsInput, "CtrlAsInput");
                this.DynamicsOff = this.AddColumn(VA.ShapeSheet.SRCConstants.DynamicsOff, "DynamicsOff");
                this.EnableGrid = this.AddColumn(VA.ShapeSheet.SRCConstants.EnableGrid, "EnableGrid");
                this.LineAdjustFrom = this.AddColumn(VA.ShapeSheet.SRCConstants.LineAdjustFrom, "LineAdjustFrom");
                this.LineAdjustTo = this.AddColumn(VA.ShapeSheet.SRCConstants.LineAdjustTo, "LineAdjustTo");
                this.LineJumpCode = this.AddColumn(VA.ShapeSheet.SRCConstants.LineJumpCode, "LineJumpCode");
                this.LineJumpFactorX = this.AddColumn(VA.ShapeSheet.SRCConstants.LineJumpFactorX, "LineJumpFactorX");
                this.LineJumpFactorY = this.AddColumn(VA.ShapeSheet.SRCConstants.LineJumpFactorY, "LineJumpFactorY");
                this.LineJumpStyle = this.AddColumn(VA.ShapeSheet.SRCConstants.LineJumpStyle, "LineJumpStyle");
                this.LineRouteExt = this.AddColumn(VA.ShapeSheet.SRCConstants.LineRouteExt, "LineRouteExt");
                this.LineToLineX = this.AddColumn(VA.ShapeSheet.SRCConstants.LineToLineX, "LineToLineX");
                this.LineToLineY = this.AddColumn(VA.ShapeSheet.SRCConstants.LineToLineY, "LineToLineY");
                this.LineToNodeX = this.AddColumn(VA.ShapeSheet.SRCConstants.LineToNodeX, "LineToNodeX");
                this.LineToNodeY = this.AddColumn(VA.ShapeSheet.SRCConstants.LineToNodeY, "LineToNodeY");
                this.PageLineJumpDirX = this.AddColumn(VA.ShapeSheet.SRCConstants.PageLineJumpDirX, "PageLineJumpDirX");
                this.PageLineJumpDirY = this.AddColumn(VA.ShapeSheet.SRCConstants.PageLineJumpDirY, "PageLineJumpDirY");
                this.PageShapeSplit = this.AddColumn(VA.ShapeSheet.SRCConstants.PageShapeSplit, "PageShapeSplit");
                this.PlaceDepth = this.AddColumn(VA.ShapeSheet.SRCConstants.PlaceDepth, "PlaceDepth");
                this.PlaceFlip = this.AddColumn(VA.ShapeSheet.SRCConstants.PlaceFlip, "PlaceFlip");
                this.PlaceStyle = this.AddColumn(VA.ShapeSheet.SRCConstants.PlaceStyle, "PlaceStyle");
                this.PlowCode = this.AddColumn(VA.ShapeSheet.SRCConstants.PlowCode, "PlowCode");
                this.ResizePage = this.AddColumn(VA.ShapeSheet.SRCConstants.ResizePage, "ResizePage");
                this.RouteStyle = this.AddColumn(VA.ShapeSheet.SRCConstants.RouteStyle, "RouteStyle");
            }

        }

    }
}