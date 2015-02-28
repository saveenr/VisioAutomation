using System.Collections.Generic;
using IVisio=Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;
using VisioAutomation.Extensions;
using VisioAutomation.ShapeSheet.Query;

namespace VisioAutomation.Pages
{
    public class PageCells : VA.ShapeSheet.CellGroups.CellGroup
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
        public VA.ShapeSheet.CellData<int> AvoidPageBreaks { get; set; } // new in visio 2010
        public VA.ShapeSheet.CellData<int> DrawingResizeType { get; set; } // new in visio 2010

        public override IEnumerable<SRCFormulaPair> Pairs
        {
            get
            {
                yield return newpair(ShapeSheet.SRCConstants.PageLeftMargin, this.PageLeftMargin.Formula);
                yield return newpair(ShapeSheet.SRCConstants.CenterX, this.CenterX.Formula);
                yield return newpair(ShapeSheet.SRCConstants.CenterY, this.CenterY.Formula);
                yield return newpair(ShapeSheet.SRCConstants.OnPage, this.OnPage.Formula);
                yield return newpair(ShapeSheet.SRCConstants.PageBottomMargin, this.PageBottomMargin.Formula);
                yield return newpair(ShapeSheet.SRCConstants.PageRightMargin, this.PageRightMargin.Formula);
                yield return newpair(ShapeSheet.SRCConstants.PagesX, this.PagesX.Formula);
                yield return newpair(ShapeSheet.SRCConstants.PagesY, this.PagesY.Formula);
                yield return newpair(ShapeSheet.SRCConstants.PageTopMargin, this.PageTopMargin.Formula);
                yield return newpair(ShapeSheet.SRCConstants.PaperKind, this.PaperKind.Formula);
                yield return newpair(ShapeSheet.SRCConstants.PrintGrid, this.PrintGrid.Formula);
                yield return newpair(ShapeSheet.SRCConstants.PrintPageOrientation, this.PrintPageOrientation.Formula);
                yield return newpair(ShapeSheet.SRCConstants.ScaleX, this.ScaleX.Formula);
                yield return newpair(ShapeSheet.SRCConstants.ScaleY, this.ScaleY.Formula);
                yield return newpair(ShapeSheet.SRCConstants.PaperSource, this.PaperSource.Formula);
                yield return newpair(ShapeSheet.SRCConstants.DrawingScale, this.DrawingScale.Formula);
                yield return newpair(ShapeSheet.SRCConstants.DrawingScaleType, this.DrawingScaleType.Formula);
                yield return newpair(ShapeSheet.SRCConstants.DrawingSizeType, this.DrawingSizeType.Formula);
                yield return newpair(ShapeSheet.SRCConstants.InhibitSnap, this.InhibitSnap.Formula);
                yield return newpair(ShapeSheet.SRCConstants.PageHeight, this.PageHeight.Formula);
                yield return newpair(ShapeSheet.SRCConstants.PageScale, this.PageScale.Formula);
                yield return newpair(ShapeSheet.SRCConstants.PageWidth, this.PageWidth.Formula);
                yield return newpair(ShapeSheet.SRCConstants.ShdwObliqueAngle, this.ShdwObliqueAngle.Formula);
                yield return newpair(ShapeSheet.SRCConstants.ShdwOffsetX, this.ShdwOffsetX.Formula);
                yield return newpair(ShapeSheet.SRCConstants.ShdwOffsetY, this.ShdwOffsetY.Formula);
                yield return newpair(ShapeSheet.SRCConstants.ShdwScaleFactor, this.ShdwScaleFactor.Formula);
                yield return newpair(ShapeSheet.SRCConstants.ShdwType, this.ShdwType.Formula);
                yield return newpair(ShapeSheet.SRCConstants.UIVisibility, this.UIVisibility.Formula);
                yield return newpair(ShapeSheet.SRCConstants.XGridDensity, this.XGridDensity.Formula);
                yield return newpair(ShapeSheet.SRCConstants.XGridOrigin, this.XGridOrigin.Formula);
                yield return newpair(ShapeSheet.SRCConstants.XGridSpacing, this.XGridSpacing.Formula);
                yield return newpair(ShapeSheet.SRCConstants.XRulerDensity, this.XRulerDensity.Formula);
                yield return newpair(ShapeSheet.SRCConstants.XRulerOrigin, this.XRulerOrigin.Formula);
                yield return newpair(ShapeSheet.SRCConstants.YGridDensity, this.YGridDensity.Formula);
                yield return newpair(ShapeSheet.SRCConstants.YGridOrigin, this.YGridOrigin.Formula);
                yield return newpair(ShapeSheet.SRCConstants.YGridSpacing, this.YGridSpacing.Formula);
                yield return newpair(ShapeSheet.SRCConstants.YRulerDensity, this.YRulerDensity.Formula);
                yield return newpair(ShapeSheet.SRCConstants.YRulerOrigin, this.YRulerOrigin.Formula);
                yield return newpair(ShapeSheet.SRCConstants.AvenueSizeX, this.AvenueSizeX.Formula);
                yield return newpair(ShapeSheet.SRCConstants.AvenueSizeY, this.AvenueSizeY.Formula);
                yield return newpair(ShapeSheet.SRCConstants.BlockSizeX, this.BlockSizeX.Formula);
                yield return newpair(ShapeSheet.SRCConstants.BlockSizeY, this.BlockSizeY.Formula);
                yield return newpair(ShapeSheet.SRCConstants.CtrlAsInput, this.CtrlAsInput.Formula);
                yield return newpair(ShapeSheet.SRCConstants.DynamicsOff, this.DynamicsOff.Formula);
                yield return newpair(ShapeSheet.SRCConstants.EnableGrid, this.EnableGrid.Formula);
                yield return newpair(ShapeSheet.SRCConstants.LineAdjustFrom, this.LineAdjustFrom.Formula);
                yield return newpair(ShapeSheet.SRCConstants.LineAdjustTo, this.LineAdjustTo.Formula);
                yield return newpair(ShapeSheet.SRCConstants.LineJumpCode, this.LineJumpCode.Formula);
                yield return newpair(ShapeSheet.SRCConstants.LineJumpFactorX, this.LineJumpFactorX.Formula);
                yield return newpair(ShapeSheet.SRCConstants.LineJumpFactorY, this.LineJumpFactorY.Formula);
                yield return newpair(ShapeSheet.SRCConstants.LineJumpStyle, this.LineJumpStyle.Formula);
                yield return newpair(ShapeSheet.SRCConstants.LineRouteExt, this.LineRouteExt.Formula);
                yield return newpair(ShapeSheet.SRCConstants.LineToLineX, this.LineToLineX.Formula);
                yield return newpair(ShapeSheet.SRCConstants.LineToLineY, this.LineToLineY.Formula);
                yield return newpair(ShapeSheet.SRCConstants.LineToNodeX, this.LineToNodeX.Formula);
                yield return newpair(ShapeSheet.SRCConstants.LineToNodeY, this.LineToNodeY.Formula);
                yield return newpair(ShapeSheet.SRCConstants.PageLineJumpDirX, this.PageLineJumpDirX.Formula);
                yield return newpair(ShapeSheet.SRCConstants.PageLineJumpDirY, this.PageLineJumpDirY.Formula);
                yield return newpair(ShapeSheet.SRCConstants.PageShapeSplit, this.PageShapeSplit.Formula);
                yield return newpair(ShapeSheet.SRCConstants.PlaceDepth, this.PlaceDepth.Formula);
                yield return newpair(ShapeSheet.SRCConstants.PlaceFlip, this.PlaceFlip.Formula);
                yield return newpair(ShapeSheet.SRCConstants.PlaceStyle, this.PlaceStyle.Formula);
                yield return newpair(ShapeSheet.SRCConstants.PlowCode, this.PlowCode.Formula);
                yield return newpair(ShapeSheet.SRCConstants.ResizePage, this.ResizePage.Formula);
                yield return newpair(ShapeSheet.SRCConstants.RouteStyle, this.RouteStyle.Formula);
                yield return newpair(ShapeSheet.SRCConstants.AvoidPageBreaks, this.AvoidPageBreaks.Formula);
                yield return newpair(ShapeSheet.SRCConstants.DrawingResizeType, this.DrawingResizeType.Formula);
            }
        }

        public static PageCells GetCells(IVisio.Shape shape)
        {
            var query = get_query();
            return VA.ShapeSheet.CellGroups.CellGroup._GetCells<PageCells, double>(shape, query, query.GetCells);
        }

        private static PageCellQuery _mCellQuery;
        private static PageCellQuery get_query()
        {
            _mCellQuery = _mCellQuery ?? new PageCellQuery();
            return _mCellQuery;
        }

        class PageCellQuery : VA.ShapeSheet.Query.CellQuery
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
                this.PageLeftMargin = this.AddCell(VA.ShapeSheet.SRCConstants.PageLeftMargin);
                this.CenterX = this.AddCell(VA.ShapeSheet.SRCConstants.CenterX);
                this.CenterY = this.AddCell(VA.ShapeSheet.SRCConstants.CenterY);
                this.OnPage = this.AddCell(VA.ShapeSheet.SRCConstants.OnPage);
                this.PageBottomMargin = this.AddCell(VA.ShapeSheet.SRCConstants.PageBottomMargin);
                this.PageRightMargin = this.AddCell(VA.ShapeSheet.SRCConstants.PageRightMargin);
                this.PagesX = this.AddCell(VA.ShapeSheet.SRCConstants.PagesX);
                this.PagesY = this.AddCell(VA.ShapeSheet.SRCConstants.PagesY);
                this.PageTopMargin = this.AddCell(VA.ShapeSheet.SRCConstants.PageTopMargin);
                this.PaperKind = this.AddCell(VA.ShapeSheet.SRCConstants.PaperKind);
                this.PrintGrid = this.AddCell(VA.ShapeSheet.SRCConstants.PrintGrid);
                this.PrintPageOrientation = this.AddCell(VA.ShapeSheet.SRCConstants.PrintPageOrientation);
                this.ScaleX = this.AddCell(VA.ShapeSheet.SRCConstants.ScaleX);
                this.ScaleY = this.AddCell(VA.ShapeSheet.SRCConstants.ScaleY);
                this.PaperSource = this.AddCell(VA.ShapeSheet.SRCConstants.PaperSource);
                this.DrawingScale = this.AddCell(VA.ShapeSheet.SRCConstants.DrawingScale);
                this.DrawingScaleType = this.AddCell(VA.ShapeSheet.SRCConstants.DrawingScaleType);
                this.DrawingSizeType = this.AddCell(VA.ShapeSheet.SRCConstants.DrawingSizeType);
                this.InhibitSnap = this.AddCell(VA.ShapeSheet.SRCConstants.InhibitSnap);
                this.PageHeight = this.AddCell(VA.ShapeSheet.SRCConstants.PageHeight);
                this.PageScale = this.AddCell(VA.ShapeSheet.SRCConstants.PageScale);
                this.PageWidth = this.AddCell(VA.ShapeSheet.SRCConstants.PageWidth);
                this.ShdwObliqueAngle = this.AddCell(VA.ShapeSheet.SRCConstants.ShdwObliqueAngle);
                this.ShdwOffsetX = this.AddCell(VA.ShapeSheet.SRCConstants.ShdwOffsetX);
                this.ShdwOffsetY = this.AddCell(VA.ShapeSheet.SRCConstants.ShdwOffsetY);
                this.ShdwScaleFactor = this.AddCell(VA.ShapeSheet.SRCConstants.ShdwScaleFactor);
                this.ShdwType = this.AddCell(VA.ShapeSheet.SRCConstants.ShdwType);
                this.UIVisibility = this.AddCell(VA.ShapeSheet.SRCConstants.UIVisibility);
                this.XGridDensity = this.AddCell(VA.ShapeSheet.SRCConstants.XGridDensity);
                this.XGridOrigin = this.AddCell(VA.ShapeSheet.SRCConstants.XGridOrigin);
                this.XGridSpacing = this.AddCell(VA.ShapeSheet.SRCConstants.XGridSpacing);
                this.XRulerDensity = this.AddCell(VA.ShapeSheet.SRCConstants.XRulerDensity);
                this.XRulerOrigin = this.AddCell(VA.ShapeSheet.SRCConstants.XRulerOrigin);
                this.YGridDensity = this.AddCell(VA.ShapeSheet.SRCConstants.YGridDensity);
                this.YGridOrigin = this.AddCell(VA.ShapeSheet.SRCConstants.YGridOrigin);
                this.YGridSpacing = this.AddCell(VA.ShapeSheet.SRCConstants.YGridSpacing);
                this.YRulerDensity = this.AddCell(VA.ShapeSheet.SRCConstants.YRulerDensity);
                this.YRulerOrigin = this.AddCell(VA.ShapeSheet.SRCConstants.YRulerOrigin);
                this.AvenueSizeX = this.AddCell(VA.ShapeSheet.SRCConstants.AvenueSizeX);
                this.AvenueSizeY = this.AddCell(VA.ShapeSheet.SRCConstants.AvenueSizeY);
                this.BlockSizeX = this.AddCell(VA.ShapeSheet.SRCConstants.BlockSizeX);
                this.BlockSizeY = this.AddCell(VA.ShapeSheet.SRCConstants.BlockSizeY);
                this.CtrlAsInput = this.AddCell(VA.ShapeSheet.SRCConstants.CtrlAsInput);
                this.DynamicsOff = this.AddCell(VA.ShapeSheet.SRCConstants.DynamicsOff);
                this.EnableGrid = this.AddCell(VA.ShapeSheet.SRCConstants.EnableGrid);
                this.LineAdjustFrom = this.AddCell(VA.ShapeSheet.SRCConstants.LineAdjustFrom);
                this.LineAdjustTo = this.AddCell(VA.ShapeSheet.SRCConstants.LineAdjustTo);
                this.LineJumpCode = this.AddCell(VA.ShapeSheet.SRCConstants.LineJumpCode);
                this.LineJumpFactorX = this.AddCell(VA.ShapeSheet.SRCConstants.LineJumpFactorX);
                this.LineJumpFactorY = this.AddCell(VA.ShapeSheet.SRCConstants.LineJumpFactorY);
                this.LineJumpStyle = this.AddCell(VA.ShapeSheet.SRCConstants.LineJumpStyle);
                this.LineRouteExt = this.AddCell(VA.ShapeSheet.SRCConstants.LineRouteExt);
                this.LineToLineX = this.AddCell(VA.ShapeSheet.SRCConstants.LineToLineX);
                this.LineToLineY = this.AddCell(VA.ShapeSheet.SRCConstants.LineToLineY);
                this.LineToNodeX = this.AddCell(VA.ShapeSheet.SRCConstants.LineToNodeX);
                this.LineToNodeY = this.AddCell(VA.ShapeSheet.SRCConstants.LineToNodeY);
                this.PageLineJumpDirX = this.AddCell(VA.ShapeSheet.SRCConstants.PageLineJumpDirX);
                this.PageLineJumpDirY = this.AddCell(VA.ShapeSheet.SRCConstants.PageLineJumpDirY);
                this.PageShapeSplit = this.AddCell(VA.ShapeSheet.SRCConstants.PageShapeSplit);
                this.PlaceDepth = this.AddCell(VA.ShapeSheet.SRCConstants.PlaceDepth);
                this.PlaceFlip = this.AddCell(VA.ShapeSheet.SRCConstants.PlaceFlip);
                this.PlaceStyle = this.AddCell(VA.ShapeSheet.SRCConstants.PlaceStyle);
                this.PlowCode = this.AddCell(VA.ShapeSheet.SRCConstants.PlowCode);
                this.ResizePage = this.AddCell(VA.ShapeSheet.SRCConstants.ResizePage);
                this.RouteStyle = this.AddCell(VA.ShapeSheet.SRCConstants.RouteStyle);
                this.AvoidPageBreaks = this.AddCell(VA.ShapeSheet.SRCConstants.AvoidPageBreaks);
                this.DrawingResizeType = this.AddCell(VA.ShapeSheet.SRCConstants.DrawingResizeType);
            }


            public PageCells GetCells(IList<VA.ShapeSheet.CellData<double>> row)
            {

                var cells = new PageCells();
                cells.PageLeftMargin = row[PageLeftMargin];
                cells.CenterX = row[CenterX];
                cells.CenterY = row[CenterY];
                cells.OnPage = row[OnPage].ToInt();
                cells.PageBottomMargin = row[PageBottomMargin];
                cells.PageRightMargin = row[PageRightMargin];
                cells.PagesX = row[PagesX];
                cells.PagesY = row[PagesY];
                cells.PageTopMargin = row[PageTopMargin];
                cells.PaperKind = row[PaperKind].ToInt();
                cells.PrintGrid = row[PrintGrid].ToInt();
                cells.PrintPageOrientation = row[PrintPageOrientation].ToInt();
                cells.ScaleX = row[ScaleX];
                cells.ScaleY = row[ScaleY];
                cells.PaperSource = row[PaperSource].ToInt();
                cells.DrawingScale = row[DrawingScale];
                cells.DrawingScaleType = row[DrawingScaleType].ToInt();
                cells.DrawingSizeType = row[DrawingSizeType].ToInt();
                cells.InhibitSnap = row[InhibitSnap].ToInt();
                cells.PageHeight = row[PageHeight];
                cells.PageScale = row[PageScale];
                cells.PageWidth = row[PageWidth];
                cells.ShdwObliqueAngle = row[ShdwObliqueAngle];
                cells.ShdwOffsetX = row[ShdwOffsetX];
                cells.ShdwOffsetY = row[ShdwOffsetY];
                cells.ShdwScaleFactor = row[ShdwScaleFactor];
                cells.ShdwType = row[ShdwType].ToInt();
                cells.UIVisibility = row[UIVisibility];
                cells.XGridDensity = row[XGridDensity];
                cells.XGridOrigin = row[XGridOrigin];
                cells.XGridSpacing = row[XGridSpacing];
                cells.XRulerDensity = row[XRulerDensity];
                cells.XRulerOrigin = row[XRulerOrigin];
                cells.YGridDensity = row[YGridDensity];
                cells.YGridOrigin = row[YGridOrigin];
                cells.YGridSpacing = row[YGridSpacing];
                cells.YRulerDensity = row[YRulerDensity];
                cells.YRulerOrigin = row[YRulerOrigin];
                cells.AvenueSizeX = row[AvenueSizeX];
                cells.AvenueSizeY = row[AvenueSizeY];
                cells.BlockSizeX = row[BlockSizeX];
                cells.BlockSizeY = row[BlockSizeY];
                cells.CtrlAsInput = row[CtrlAsInput].ToInt();
                cells.DynamicsOff = row[DynamicsOff].ToInt();
                cells.EnableGrid = row[EnableGrid].ToInt();
                cells.LineAdjustFrom = row[LineAdjustFrom].ToInt();
                cells.LineAdjustTo = row[LineAdjustTo];
                cells.LineJumpCode = row[LineJumpCode];
                cells.LineJumpFactorX = row[LineJumpFactorX];
                cells.LineJumpFactorY = row[LineJumpFactorY];
                cells.LineJumpStyle = row[LineJumpStyle].ToInt();
                cells.LineRouteExt = row[LineRouteExt];
                cells.LineToLineX = row[LineToLineX];
                cells.LineToLineY = row[LineToLineY];
                cells.LineToNodeX = row[LineToNodeX];
                cells.LineToNodeY = row[LineToNodeY];
                cells.PageLineJumpDirX = row[PageLineJumpDirX];
                cells.PageLineJumpDirY = row[PageLineJumpDirY];
                cells.PageShapeSplit = row[PageShapeSplit].ToInt();
                cells.PlaceDepth = row[PlaceDepth].ToInt();
                cells.PlaceFlip = row[PlaceFlip].ToInt();
                cells.PlaceStyle = row[PlaceStyle].ToInt();
                cells.PlowCode = row[PlowCode].ToInt();
                cells.ResizePage = row[ResizePage].ToInt();
                cells.RouteStyle = row[RouteStyle].ToInt();
                cells.AvoidPageBreaks = row[AvoidPageBreaks].ToInt();
                cells.DrawingResizeType = row[DrawingResizeType].ToInt();
                return cells;
            }

        }

    }
}