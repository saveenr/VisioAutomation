using System.Collections.Generic;
using IVisio=Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;
using VisioAutomation.Extensions;
using VisioAutomation.ShapeSheet.Query;

namespace VisioAutomation.Pages
{
    public class PageCells : ShapeSheet.CellGroups.CellGroup
    {
        public ShapeSheet.CellData<double> PageLeftMargin { get; set; }
        public ShapeSheet.CellData<double> CenterX { get; set; }
        public ShapeSheet.CellData<double> CenterY { get; set; }
        public ShapeSheet.CellData<int> OnPage { get; set; }
        public ShapeSheet.CellData<double> PageBottomMargin { get; set; }
        public ShapeSheet.CellData<double> PageRightMargin { get; set; }
        public ShapeSheet.CellData<double> PagesX { get; set; }
        public ShapeSheet.CellData<double> PagesY { get; set; }
        public ShapeSheet.CellData<double> PageTopMargin { get; set; }
        public ShapeSheet.CellData<int> PaperKind { get; set; }
        public ShapeSheet.CellData<int> PrintGrid { get; set; }
        public ShapeSheet.CellData<int> PrintPageOrientation { get; set; }
        public ShapeSheet.CellData<double> ScaleX { get; set; }
        public ShapeSheet.CellData<double> ScaleY { get; set; }
        public ShapeSheet.CellData<int> PaperSource { get; set; }
        public ShapeSheet.CellData<double> DrawingScale { get; set; }
        public ShapeSheet.CellData<int> DrawingScaleType { get; set; }
        public ShapeSheet.CellData<int> DrawingSizeType { get; set; }
        public ShapeSheet.CellData<int> InhibitSnap { get; set; }
        public ShapeSheet.CellData<double> PageHeight { get; set; }
        public ShapeSheet.CellData<double> PageScale { get; set; }
        public ShapeSheet.CellData<double> PageWidth { get; set; }
        public ShapeSheet.CellData<double> ShdwObliqueAngle { get; set; }
        public ShapeSheet.CellData<double> ShdwOffsetX { get; set; }
        public ShapeSheet.CellData<double> ShdwOffsetY { get; set; }
        public ShapeSheet.CellData<double> ShdwScaleFactor { get; set; }
        public ShapeSheet.CellData<int> ShdwType { get; set; }
        public ShapeSheet.CellData<double> UIVisibility { get; set; }
        public ShapeSheet.CellData<double> XGridDensity { get; set; }
        public ShapeSheet.CellData<double> XGridOrigin { get; set; }
        public ShapeSheet.CellData<double> XGridSpacing { get; set; }
        public ShapeSheet.CellData<double> XRulerDensity { get; set; }
        public ShapeSheet.CellData<double> XRulerOrigin { get; set; }
        public ShapeSheet.CellData<double> YGridDensity { get; set; }
        public ShapeSheet.CellData<double> YGridOrigin { get; set; }
        public ShapeSheet.CellData<double> YGridSpacing { get; set; }
        public ShapeSheet.CellData<double> YRulerDensity { get; set; }
        public ShapeSheet.CellData<double> YRulerOrigin { get; set; }
        public ShapeSheet.CellData<double> AvenueSizeX { get; set; }
        public ShapeSheet.CellData<double> AvenueSizeY { get; set; }
        public ShapeSheet.CellData<double> BlockSizeX { get; set; }
        public ShapeSheet.CellData<double> BlockSizeY { get; set; }
        public ShapeSheet.CellData<int> CtrlAsInput { get; set; }
        public ShapeSheet.CellData<int> DynamicsOff { get; set; }
        public ShapeSheet.CellData<int> EnableGrid { get; set; }
        public ShapeSheet.CellData<int> LineAdjustFrom { get; set; }
        public ShapeSheet.CellData<double> LineAdjustTo { get; set; }
        public ShapeSheet.CellData<double> LineJumpCode { get; set; }
        public ShapeSheet.CellData<double> LineJumpFactorX { get; set; }
        public ShapeSheet.CellData<double> LineJumpFactorY { get; set; }
        public ShapeSheet.CellData<int> LineJumpStyle { get; set; }
        public ShapeSheet.CellData<double> LineRouteExt { get; set; }
        public ShapeSheet.CellData<double> LineToLineX { get; set; }
        public ShapeSheet.CellData<double> LineToLineY { get; set; }
        public ShapeSheet.CellData<double> LineToNodeX { get; set; }
        public ShapeSheet.CellData<double> LineToNodeY { get; set; }
        public ShapeSheet.CellData<double> PageLineJumpDirX { get; set; }
        public ShapeSheet.CellData<double> PageLineJumpDirY { get; set; }
        public ShapeSheet.CellData<int> PageShapeSplit { get; set; }
        public ShapeSheet.CellData<int> PlaceDepth { get; set; }
        public ShapeSheet.CellData<int> PlaceFlip { get; set; }
        public ShapeSheet.CellData<int> PlaceStyle { get; set; }
        public ShapeSheet.CellData<int> PlowCode { get; set; }
        public ShapeSheet.CellData<int> ResizePage { get; set; }
        public ShapeSheet.CellData<int> RouteStyle { get; set; }
        public ShapeSheet.CellData<int> AvoidPageBreaks { get; set; } // new in visio 2010
        public ShapeSheet.CellData<int> DrawingResizeType { get; set; } // new in visio 2010

        public override IEnumerable<SRCFormulaPair> Pairs
        {
            get
            {
                yield return this.newpair(ShapeSheet.SRCConstants.PageLeftMargin, this.PageLeftMargin.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.CenterX, this.CenterX.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.CenterY, this.CenterY.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.OnPage, this.OnPage.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.PageBottomMargin, this.PageBottomMargin.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.PageRightMargin, this.PageRightMargin.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.PagesX, this.PagesX.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.PagesY, this.PagesY.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.PageTopMargin, this.PageTopMargin.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.PaperKind, this.PaperKind.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.PrintGrid, this.PrintGrid.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.PrintPageOrientation, this.PrintPageOrientation.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.ScaleX, this.ScaleX.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.ScaleY, this.ScaleY.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.PaperSource, this.PaperSource.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.DrawingScale, this.DrawingScale.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.DrawingScaleType, this.DrawingScaleType.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.DrawingSizeType, this.DrawingSizeType.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.InhibitSnap, this.InhibitSnap.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.PageHeight, this.PageHeight.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.PageScale, this.PageScale.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.PageWidth, this.PageWidth.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.ShdwObliqueAngle, this.ShdwObliqueAngle.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.ShdwOffsetX, this.ShdwOffsetX.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.ShdwOffsetY, this.ShdwOffsetY.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.ShdwScaleFactor, this.ShdwScaleFactor.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.ShdwType, this.ShdwType.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.UIVisibility, this.UIVisibility.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.XGridDensity, this.XGridDensity.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.XGridOrigin, this.XGridOrigin.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.XGridSpacing, this.XGridSpacing.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.XRulerDensity, this.XRulerDensity.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.XRulerOrigin, this.XRulerOrigin.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.YGridDensity, this.YGridDensity.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.YGridOrigin, this.YGridOrigin.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.YGridSpacing, this.YGridSpacing.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.YRulerDensity, this.YRulerDensity.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.YRulerOrigin, this.YRulerOrigin.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.AvenueSizeX, this.AvenueSizeX.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.AvenueSizeY, this.AvenueSizeY.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.BlockSizeX, this.BlockSizeX.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.BlockSizeY, this.BlockSizeY.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.CtrlAsInput, this.CtrlAsInput.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.DynamicsOff, this.DynamicsOff.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.EnableGrid, this.EnableGrid.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.LineAdjustFrom, this.LineAdjustFrom.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.LineAdjustTo, this.LineAdjustTo.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.LineJumpCode, this.LineJumpCode.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.LineJumpFactorX, this.LineJumpFactorX.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.LineJumpFactorY, this.LineJumpFactorY.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.LineJumpStyle, this.LineJumpStyle.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.LineRouteExt, this.LineRouteExt.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.LineToLineX, this.LineToLineX.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.LineToLineY, this.LineToLineY.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.LineToNodeX, this.LineToNodeX.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.LineToNodeY, this.LineToNodeY.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.PageLineJumpDirX, this.PageLineJumpDirX.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.PageLineJumpDirY, this.PageLineJumpDirY.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.PageShapeSplit, this.PageShapeSplit.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.PlaceDepth, this.PlaceDepth.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.PlaceFlip, this.PlaceFlip.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.PlaceStyle, this.PlaceStyle.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.PlowCode, this.PlowCode.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.ResizePage, this.ResizePage.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.RouteStyle, this.RouteStyle.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.AvoidPageBreaks, this.AvoidPageBreaks.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.DrawingResizeType, this.DrawingResizeType.Formula);
            }
        }

        public static PageCells GetCells(IVisio.Shape shape)
        {
            var query = get_query();
            return _GetCells<PageCells, double>(shape, query, query.GetCells);
        }

        private static PageCellQuery _mCellQuery;
        private static PageCellQuery get_query()
        {
            _mCellQuery = _mCellQuery ?? new PageCellQuery();
            return _mCellQuery;
        }

        class PageCellQuery : CellQuery
        {
            public CellColumn PageLeftMargin { get; set; }
            public CellColumn CenterX { get; set; }
            public CellColumn CenterY { get; set; }
            public CellColumn OnPage { get; set; }
            public CellColumn PageBottomMargin { get; set; }
            public CellColumn PageRightMargin { get; set; }
            public CellColumn PagesX { get; set; }
            public CellColumn PagesY { get; set; }
            public CellColumn PageTopMargin { get; set; }
            public CellColumn PaperKind { get; set; }
            public CellColumn PrintGrid { get; set; }
            public CellColumn PrintPageOrientation { get; set; }
            public CellColumn ScaleX { get; set; }
            public CellColumn ScaleY { get; set; }
            public CellColumn PaperSource { get; set; }
            public CellColumn DrawingScale { get; set; }
            public CellColumn DrawingScaleType { get; set; }
            public CellColumn DrawingSizeType { get; set; }
            public CellColumn InhibitSnap { get; set; }
            public CellColumn PageHeight { get; set; }
            public CellColumn PageScale { get; set; }
            public CellColumn PageWidth { get; set; }
            public CellColumn ShdwObliqueAngle { get; set; }
            public CellColumn ShdwOffsetX { get; set; }
            public CellColumn ShdwOffsetY { get; set; }
            public CellColumn ShdwScaleFactor { get; set; }
            public CellColumn ShdwType { get; set; }
            public CellColumn UIVisibility { get; set; }
            public CellColumn XGridDensity { get; set; }
            public CellColumn XGridOrigin { get; set; }
            public CellColumn XGridSpacing { get; set; }
            public CellColumn XRulerDensity { get; set; }
            public CellColumn XRulerOrigin { get; set; }
            public CellColumn YGridDensity { get; set; }
            public CellColumn YGridOrigin { get; set; }
            public CellColumn YGridSpacing { get; set; }
            public CellColumn YRulerDensity { get; set; }
            public CellColumn YRulerOrigin { get; set; }
            public CellColumn AvenueSizeX { get; set; }
            public CellColumn AvenueSizeY { get; set; }
            public CellColumn BlockSizeX { get; set; }
            public CellColumn BlockSizeY { get; set; }
            public CellColumn CtrlAsInput { get; set; }
            public CellColumn DynamicsOff { get; set; }
            public CellColumn EnableGrid { get; set; }
            public CellColumn LineAdjustFrom { get; set; }
            public CellColumn LineAdjustTo { get; set; }
            public CellColumn LineJumpCode { get; set; }
            public CellColumn LineJumpFactorX { get; set; }
            public CellColumn LineJumpFactorY { get; set; }
            public CellColumn LineJumpStyle { get; set; }
            public CellColumn LineRouteExt { get; set; }
            public CellColumn LineToLineX { get; set; }
            public CellColumn LineToLineY { get; set; }
            public CellColumn LineToNodeX { get; set; }
            public CellColumn LineToNodeY { get; set; }
            public CellColumn PageLineJumpDirX { get; set; }
            public CellColumn PageLineJumpDirY { get; set; }
            public CellColumn PageShapeSplit { get; set; }
            public CellColumn PlaceDepth { get; set; }
            public CellColumn PlaceFlip { get; set; }
            public CellColumn PlaceStyle { get; set; }
            public CellColumn PlowCode { get; set; }
            public CellColumn ResizePage { get; set; }
            public CellColumn RouteStyle { get; set; }
            public CellColumn AvoidPageBreaks { get; set; }
            public CellColumn DrawingResizeType { get; set; }

            public PageCellQuery() 
            {
                this.PageLeftMargin = this.AddCell(ShapeSheet.SRCConstants.PageLeftMargin, "PageLeftMargin");
                this.CenterX = this.AddCell(ShapeSheet.SRCConstants.CenterX, "CenterX");
                this.CenterY = this.AddCell(ShapeSheet.SRCConstants.CenterY, "CenterY");
                this.OnPage = this.AddCell(ShapeSheet.SRCConstants.OnPage, "OnPage");
                this.PageBottomMargin = this.AddCell(ShapeSheet.SRCConstants.PageBottomMargin, "PageBottomMargin");
                this.PageRightMargin = this.AddCell(ShapeSheet.SRCConstants.PageRightMargin, "PageRightMargin");
                this.PagesX = this.AddCell(ShapeSheet.SRCConstants.PagesX, "PagesX");
                this.PagesY = this.AddCell(ShapeSheet.SRCConstants.PagesY, "PagesY");
                this.PageTopMargin = this.AddCell(ShapeSheet.SRCConstants.PageTopMargin, "PageTopMargin");
                this.PaperKind = this.AddCell(ShapeSheet.SRCConstants.PaperKind, "PaperKind");
                this.PrintGrid = this.AddCell(ShapeSheet.SRCConstants.PrintGrid, "PrintGrid");
                this.PrintPageOrientation = this.AddCell(ShapeSheet.SRCConstants.PrintPageOrientation, "PrintPageOrientation");
                this.ScaleX = this.AddCell(ShapeSheet.SRCConstants.ScaleX, "ScaleX");
                this.ScaleY = this.AddCell(ShapeSheet.SRCConstants.ScaleY, "ScaleY");
                this.PaperSource = this.AddCell(ShapeSheet.SRCConstants.PaperSource, "PaperSource");
                this.DrawingScale = this.AddCell(ShapeSheet.SRCConstants.DrawingScale, "DrawingScale");
                this.DrawingScaleType = this.AddCell(ShapeSheet.SRCConstants.DrawingScaleType, "DrawingScaleType");
                this.DrawingSizeType = this.AddCell(ShapeSheet.SRCConstants.DrawingSizeType, "DrawingSizeType");
                this.InhibitSnap = this.AddCell(ShapeSheet.SRCConstants.InhibitSnap, "InhibitSnap");
                this.PageHeight = this.AddCell(ShapeSheet.SRCConstants.PageHeight, "PageHeight");
                this.PageScale = this.AddCell(ShapeSheet.SRCConstants.PageScale, "PageScale");
                this.PageWidth = this.AddCell(ShapeSheet.SRCConstants.PageWidth, "PageWidth");
                this.ShdwObliqueAngle = this.AddCell(ShapeSheet.SRCConstants.ShdwObliqueAngle, "ShdwObliqueAngle");
                this.ShdwOffsetX = this.AddCell(ShapeSheet.SRCConstants.ShdwOffsetX, "ShdwOffsetX");
                this.ShdwOffsetY = this.AddCell(ShapeSheet.SRCConstants.ShdwOffsetY, "ShdwOffsetY");
                this.ShdwScaleFactor = this.AddCell(ShapeSheet.SRCConstants.ShdwScaleFactor, "ShdwScaleFactor");
                this.ShdwType = this.AddCell(ShapeSheet.SRCConstants.ShdwType, "ShdwType");
                this.UIVisibility = this.AddCell(ShapeSheet.SRCConstants.UIVisibility, "UIVisibility");
                this.XGridDensity = this.AddCell(ShapeSheet.SRCConstants.XGridDensity, "XGridDensity");
                this.XGridOrigin = this.AddCell(ShapeSheet.SRCConstants.XGridOrigin, "XGridOrigin");
                this.XGridSpacing = this.AddCell(ShapeSheet.SRCConstants.XGridSpacing, "XGridSpacing");
                this.XRulerDensity = this.AddCell(ShapeSheet.SRCConstants.XRulerDensity, "XRulerDensity");
                this.XRulerOrigin = this.AddCell(ShapeSheet.SRCConstants.XRulerOrigin, "XRulerOrigin");
                this.YGridDensity = this.AddCell(ShapeSheet.SRCConstants.YGridDensity, "YGridDensity");
                this.YGridOrigin = this.AddCell(ShapeSheet.SRCConstants.YGridOrigin, "YGridOrigin");
                this.YGridSpacing = this.AddCell(ShapeSheet.SRCConstants.YGridSpacing, "YGridSpacing");
                this.YRulerDensity = this.AddCell(ShapeSheet.SRCConstants.YRulerDensity, "YRulerDensity");
                this.YRulerOrigin = this.AddCell(ShapeSheet.SRCConstants.YRulerOrigin, "YRulerOrigin");
                this.AvenueSizeX = this.AddCell(ShapeSheet.SRCConstants.AvenueSizeX, "AvenueSizeX");
                this.AvenueSizeY = this.AddCell(ShapeSheet.SRCConstants.AvenueSizeY, "AvenueSizeY");
                this.BlockSizeX = this.AddCell(ShapeSheet.SRCConstants.BlockSizeX, "BlockSizeX");
                this.BlockSizeY = this.AddCell(ShapeSheet.SRCConstants.BlockSizeY, "BlockSizeY");
                this.CtrlAsInput = this.AddCell(ShapeSheet.SRCConstants.CtrlAsInput, "CtrlAsInput");
                this.DynamicsOff = this.AddCell(ShapeSheet.SRCConstants.DynamicsOff, "DynamicsOff");
                this.EnableGrid = this.AddCell(ShapeSheet.SRCConstants.EnableGrid, "EnableGrid");
                this.LineAdjustFrom = this.AddCell(ShapeSheet.SRCConstants.LineAdjustFrom, "LineAdjustFrom");
                this.LineAdjustTo = this.AddCell(ShapeSheet.SRCConstants.LineAdjustTo, "LineAdjustTo");
                this.LineJumpCode = this.AddCell(ShapeSheet.SRCConstants.LineJumpCode, "LineJumpCode");
                this.LineJumpFactorX = this.AddCell(ShapeSheet.SRCConstants.LineJumpFactorX, "LineJumpFactorX");
                this.LineJumpFactorY = this.AddCell(ShapeSheet.SRCConstants.LineJumpFactorY, "LineJumpFactorY");
                this.LineJumpStyle = this.AddCell(ShapeSheet.SRCConstants.LineJumpStyle, "LineJumpStyle");
                this.LineRouteExt = this.AddCell(ShapeSheet.SRCConstants.LineRouteExt, "LineRouteExt");
                this.LineToLineX = this.AddCell(ShapeSheet.SRCConstants.LineToLineX, "LineToLineX");
                this.LineToLineY = this.AddCell(ShapeSheet.SRCConstants.LineToLineY, "LineToLineY");
                this.LineToNodeX = this.AddCell(ShapeSheet.SRCConstants.LineToNodeX, "LineToNodeX");
                this.LineToNodeY = this.AddCell(ShapeSheet.SRCConstants.LineToNodeY, "LineToNodeY");
                this.PageLineJumpDirX = this.AddCell(ShapeSheet.SRCConstants.PageLineJumpDirX, "PageLineJumpDirX");
                this.PageLineJumpDirY = this.AddCell(ShapeSheet.SRCConstants.PageLineJumpDirY, "PageLineJumpDirY");
                this.PageShapeSplit = this.AddCell(ShapeSheet.SRCConstants.PageShapeSplit, "PageShapeSplit");
                this.PlaceDepth = this.AddCell(ShapeSheet.SRCConstants.PlaceDepth, "PlaceDepth");
                this.PlaceFlip = this.AddCell(ShapeSheet.SRCConstants.PlaceFlip, "PlaceFlip");
                this.PlaceStyle = this.AddCell(ShapeSheet.SRCConstants.PlaceStyle, "PlaceStyle");
                this.PlowCode = this.AddCell(ShapeSheet.SRCConstants.PlowCode, "PlowCode");
                this.ResizePage = this.AddCell(ShapeSheet.SRCConstants.ResizePage, "ResizePage");
                this.RouteStyle = this.AddCell(ShapeSheet.SRCConstants.RouteStyle, "RouteStyle");
                this.AvoidPageBreaks = this.AddCell(ShapeSheet.SRCConstants.AvoidPageBreaks, "AvoidPageBreaks");
                this.DrawingResizeType = this.AddCell(ShapeSheet.SRCConstants.DrawingResizeType, "DrawingResizeType");
            }


            public PageCells GetCells(IList<ShapeSheet.CellData<double>> row)
            {

                var cells = new PageCells();
                cells.PageLeftMargin = row[this.PageLeftMargin];
                cells.CenterX = row[this.CenterX];
                cells.CenterY = row[this.CenterY];
                cells.OnPage = row[this.OnPage].ToInt();
                cells.PageBottomMargin = row[this.PageBottomMargin];
                cells.PageRightMargin = row[this.PageRightMargin];
                cells.PagesX = row[this.PagesX];
                cells.PagesY = row[this.PagesY];
                cells.PageTopMargin = row[this.PageTopMargin];
                cells.PaperKind = row[this.PaperKind].ToInt();
                cells.PrintGrid = row[this.PrintGrid].ToInt();
                cells.PrintPageOrientation = row[this.PrintPageOrientation].ToInt();
                cells.ScaleX = row[this.ScaleX];
                cells.ScaleY = row[this.ScaleY];
                cells.PaperSource = row[this.PaperSource].ToInt();
                cells.DrawingScale = row[this.DrawingScale];
                cells.DrawingScaleType = row[this.DrawingScaleType].ToInt();
                cells.DrawingSizeType = row[this.DrawingSizeType].ToInt();
                cells.InhibitSnap = row[this.InhibitSnap].ToInt();
                cells.PageHeight = row[this.PageHeight];
                cells.PageScale = row[this.PageScale];
                cells.PageWidth = row[this.PageWidth];
                cells.ShdwObliqueAngle = row[this.ShdwObliqueAngle];
                cells.ShdwOffsetX = row[this.ShdwOffsetX];
                cells.ShdwOffsetY = row[this.ShdwOffsetY];
                cells.ShdwScaleFactor = row[this.ShdwScaleFactor];
                cells.ShdwType = row[this.ShdwType].ToInt();
                cells.UIVisibility = row[this.UIVisibility];
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
                cells.AvenueSizeX = row[this.AvenueSizeX];
                cells.AvenueSizeY = row[this.AvenueSizeY];
                cells.BlockSizeX = row[this.BlockSizeX];
                cells.BlockSizeY = row[this.BlockSizeY];
                cells.CtrlAsInput = row[this.CtrlAsInput].ToInt();
                cells.DynamicsOff = row[this.DynamicsOff].ToInt();
                cells.EnableGrid = row[this.EnableGrid].ToInt();
                cells.LineAdjustFrom = row[this.LineAdjustFrom].ToInt();
                cells.LineAdjustTo = row[this.LineAdjustTo];
                cells.LineJumpCode = row[this.LineJumpCode];
                cells.LineJumpFactorX = row[this.LineJumpFactorX];
                cells.LineJumpFactorY = row[this.LineJumpFactorY];
                cells.LineJumpStyle = row[this.LineJumpStyle].ToInt();
                cells.LineRouteExt = row[this.LineRouteExt];
                cells.LineToLineX = row[this.LineToLineX];
                cells.LineToLineY = row[this.LineToLineY];
                cells.LineToNodeX = row[this.LineToNodeX];
                cells.LineToNodeY = row[this.LineToNodeY];
                cells.PageLineJumpDirX = row[this.PageLineJumpDirX];
                cells.PageLineJumpDirY = row[this.PageLineJumpDirY];
                cells.PageShapeSplit = row[this.PageShapeSplit].ToInt();
                cells.PlaceDepth = row[this.PlaceDepth].ToInt();
                cells.PlaceFlip = row[this.PlaceFlip].ToInt();
                cells.PlaceStyle = row[this.PlaceStyle].ToInt();
                cells.PlowCode = row[this.PlowCode].ToInt();
                cells.ResizePage = row[this.ResizePage].ToInt();
                cells.RouteStyle = row[this.RouteStyle].ToInt();
                cells.AvoidPageBreaks = row[this.AvoidPageBreaks].ToInt();
                cells.DrawingResizeType = row[this.DrawingResizeType].ToInt();
                return cells;
            }

        }

    }
}