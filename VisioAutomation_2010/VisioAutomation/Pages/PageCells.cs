using System.Collections.Generic;
using IVisio=Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;
using VisioAutomation.Extensions;

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

        public override IEnumerable<VA.ShapeSheet.CellGroups.BaseCellGroup.SRCValuePair> EnumPairs()
        {
            yield return createpair(ShapeSheet.SRCConstants.PageLeftMargin, this.PageLeftMargin.Formula);
            yield return createpair(ShapeSheet.SRCConstants.CenterX, this.CenterX.Formula);
            yield return createpair(ShapeSheet.SRCConstants.CenterY, this.CenterY.Formula);
            yield return createpair(ShapeSheet.SRCConstants.OnPage, this.OnPage.Formula);
            yield return createpair(ShapeSheet.SRCConstants.PageBottomMargin, this.PageBottomMargin.Formula);
            yield return createpair(ShapeSheet.SRCConstants.PageRightMargin, this.PageRightMargin.Formula);
            yield return createpair(ShapeSheet.SRCConstants.PagesX, this.PagesX.Formula);
            yield return createpair(ShapeSheet.SRCConstants.PagesY, this.PagesY.Formula);
            yield return createpair(ShapeSheet.SRCConstants.PageTopMargin, this.PageTopMargin.Formula);
            yield return createpair(ShapeSheet.SRCConstants.PaperKind, this.PaperKind.Formula);
            yield return createpair(ShapeSheet.SRCConstants.PrintGrid, this.PrintGrid.Formula);
            yield return createpair(ShapeSheet.SRCConstants.PrintPageOrientation, this.PrintPageOrientation.Formula);
            yield return createpair(ShapeSheet.SRCConstants.ScaleX, this.ScaleX.Formula);
            yield return createpair(ShapeSheet.SRCConstants.ScaleY, this.ScaleY.Formula);
            yield return createpair(ShapeSheet.SRCConstants.PaperSource, this.PaperSource.Formula);
            yield return createpair(ShapeSheet.SRCConstants.DrawingScale, this.DrawingScale.Formula);
            yield return createpair(ShapeSheet.SRCConstants.DrawingScaleType, this.DrawingScaleType.Formula);
            yield return createpair(ShapeSheet.SRCConstants.DrawingSizeType, this.DrawingSizeType.Formula);
            yield return createpair(ShapeSheet.SRCConstants.InhibitSnap, this.InhibitSnap.Formula);
            yield return createpair(ShapeSheet.SRCConstants.PageHeight, this.PageHeight.Formula);
            yield return createpair(ShapeSheet.SRCConstants.PageScale, this.PageScale.Formula);
            yield return createpair(ShapeSheet.SRCConstants.PageWidth, this.PageWidth.Formula);
            yield return createpair(ShapeSheet.SRCConstants.ShdwObliqueAngle, this.ShdwObliqueAngle.Formula);
            yield return createpair(ShapeSheet.SRCConstants.ShdwOffsetX, this.ShdwOffsetX.Formula);
            yield return createpair(ShapeSheet.SRCConstants.ShdwOffsetY, this.ShdwOffsetY.Formula);
            yield return createpair(ShapeSheet.SRCConstants.ShdwScaleFactor, this.ShdwScaleFactor.Formula);
            yield return createpair(ShapeSheet.SRCConstants.ShdwType, this.ShdwType.Formula);
            yield return createpair(ShapeSheet.SRCConstants.UIVisibility, this.UIVisibility.Formula);
            yield return createpair(ShapeSheet.SRCConstants.XGridDensity, this.XGridDensity.Formula);
            yield return createpair(ShapeSheet.SRCConstants.XGridOrigin, this.XGridOrigin.Formula);
            yield return createpair(ShapeSheet.SRCConstants.XGridSpacing, this.XGridSpacing.Formula);
            yield return createpair(ShapeSheet.SRCConstants.XRulerDensity, this.XRulerDensity.Formula);
            yield return createpair(ShapeSheet.SRCConstants.XRulerOrigin, this.XRulerOrigin.Formula);
            yield return createpair(ShapeSheet.SRCConstants.YGridDensity, this.YGridDensity.Formula);
            yield return createpair(ShapeSheet.SRCConstants.YGridOrigin, this.YGridOrigin.Formula);
            yield return createpair(ShapeSheet.SRCConstants.YGridSpacing, this.YGridSpacing.Formula);
            yield return createpair(ShapeSheet.SRCConstants.YRulerDensity, this.YRulerDensity.Formula);
            yield return createpair(ShapeSheet.SRCConstants.YRulerOrigin, this.YRulerOrigin.Formula);
            yield return createpair(ShapeSheet.SRCConstants.AvenueSizeX, this.AvenueSizeX.Formula);
            yield return createpair(ShapeSheet.SRCConstants.AvenueSizeY, this.AvenueSizeY.Formula);
            yield return createpair(ShapeSheet.SRCConstants.BlockSizeX, this.BlockSizeX.Formula);
            yield return createpair(ShapeSheet.SRCConstants.BlockSizeY, this.BlockSizeY.Formula);
            yield return createpair(ShapeSheet.SRCConstants.CtrlAsInput, this.CtrlAsInput.Formula);
            yield return createpair(ShapeSheet.SRCConstants.DynamicsOff, this.DynamicsOff.Formula);
            yield return createpair(ShapeSheet.SRCConstants.EnableGrid, this.EnableGrid.Formula);
            yield return createpair(ShapeSheet.SRCConstants.LineAdjustFrom, this.LineAdjustFrom.Formula);
            yield return createpair(ShapeSheet.SRCConstants.LineAdjustTo, this.LineAdjustTo.Formula);
            yield return createpair(ShapeSheet.SRCConstants.LineJumpCode, this.LineJumpCode.Formula);
            yield return createpair(ShapeSheet.SRCConstants.LineJumpFactorX, this.LineJumpFactorX.Formula);
            yield return createpair(ShapeSheet.SRCConstants.LineJumpFactorY, this.LineJumpFactorY.Formula);
            yield return createpair(ShapeSheet.SRCConstants.LineJumpStyle, this.LineJumpStyle.Formula);
            yield return createpair(ShapeSheet.SRCConstants.LineRouteExt, this.LineRouteExt.Formula);
            yield return createpair(ShapeSheet.SRCConstants.LineToLineX, this.LineToLineX.Formula);
            yield return createpair(ShapeSheet.SRCConstants.LineToLineY, this.LineToLineY.Formula);
            yield return createpair(ShapeSheet.SRCConstants.LineToNodeX, this.LineToNodeX.Formula);
            yield return createpair(ShapeSheet.SRCConstants.LineToNodeY, this.LineToNodeY.Formula);
            yield return createpair(ShapeSheet.SRCConstants.PageLineJumpDirX, this.PageLineJumpDirX.Formula);
            yield return createpair(ShapeSheet.SRCConstants.PageLineJumpDirY, this.PageLineJumpDirY.Formula);
            yield return createpair(ShapeSheet.SRCConstants.PageShapeSplit, this.PageShapeSplit.Formula);
            yield return createpair(ShapeSheet.SRCConstants.PlaceDepth, this.PlaceDepth.Formula);
            yield return createpair(ShapeSheet.SRCConstants.PlaceFlip, this.PlaceFlip.Formula);
            yield return createpair(ShapeSheet.SRCConstants.PlaceStyle, this.PlaceStyle.Formula);
            yield return createpair(ShapeSheet.SRCConstants.PlowCode, this.PlowCode.Formula);
            yield return createpair(ShapeSheet.SRCConstants.ResizePage, this.ResizePage.Formula);
            yield return createpair(ShapeSheet.SRCConstants.RouteStyle, this.RouteStyle.Formula);
            yield return createpair(ShapeSheet.SRCConstants.AvoidPageBreaks, this.AvoidPageBreaks.Formula);
            yield return createpair(ShapeSheet.SRCConstants.DrawingResizeType, this.DrawingResizeType.Formula);
        }

        public static PageCells GetCells(IVisio.Shape shape)
        {
            var query = get_query();
            return VA.ShapeSheet.CellGroups.CellGroup._GetCells(shape, query, query.GetCells);
        }

        private static PageCellQuery _mCellQuery;
        private static PageCellQuery get_query()
        {
            _mCellQuery = _mCellQuery ?? new PageCellQuery();
            return _mCellQuery;
        }

        class PageCellQuery : VA.ShapeSheet.Query.CellQuery
        {
            public Column PageLeftMargin { get; set; }
            public Column CenterX { get; set; }
            public Column CenterY { get; set; }
            public Column OnPage { get; set; }
            public Column PageBottomMargin { get; set; }
            public Column PageRightMargin { get; set; }
            public Column PagesX { get; set; }
            public Column PagesY { get; set; }
            public Column PageTopMargin { get; set; }
            public Column PaperKind { get; set; }
            public Column PrintGrid { get; set; }
            public Column PrintPageOrientation { get; set; }
            public Column ScaleX { get; set; }
            public Column ScaleY { get; set; }
            public Column PaperSource { get; set; }
            public Column DrawingScale { get; set; }
            public Column DrawingScaleType { get; set; }
            public Column DrawingSizeType { get; set; }
            public Column InhibitSnap { get; set; }
            public Column PageHeight { get; set; }
            public Column PageScale { get; set; }
            public Column PageWidth { get; set; }
            public Column ShdwObliqueAngle { get; set; }
            public Column ShdwOffsetX { get; set; }
            public Column ShdwOffsetY { get; set; }
            public Column ShdwScaleFactor { get; set; }
            public Column ShdwType { get; set; }
            public Column UIVisibility { get; set; }
            public Column XGridDensity { get; set; }
            public Column XGridOrigin { get; set; }
            public Column XGridSpacing { get; set; }
            public Column XRulerDensity { get; set; }
            public Column XRulerOrigin { get; set; }
            public Column YGridDensity { get; set; }
            public Column YGridOrigin { get; set; }
            public Column YGridSpacing { get; set; }
            public Column YRulerDensity { get; set; }
            public Column YRulerOrigin { get; set; }
            public Column AvenueSizeX { get; set; }
            public Column AvenueSizeY { get; set; }
            public Column BlockSizeX { get; set; }
            public Column BlockSizeY { get; set; }
            public Column CtrlAsInput { get; set; }
            public Column DynamicsOff { get; set; }
            public Column EnableGrid { get; set; }
            public Column LineAdjustFrom { get; set; }
            public Column LineAdjustTo { get; set; }
            public Column LineJumpCode { get; set; }
            public Column LineJumpFactorX { get; set; }
            public Column LineJumpFactorY { get; set; }
            public Column LineJumpStyle { get; set; }
            public Column LineRouteExt { get; set; }
            public Column LineToLineX { get; set; }
            public Column LineToLineY { get; set; }
            public Column LineToNodeX { get; set; }
            public Column LineToNodeY { get; set; }
            public Column PageLineJumpDirX { get; set; }
            public Column PageLineJumpDirY { get; set; }
            public Column PageShapeSplit { get; set; }
            public Column PlaceDepth { get; set; }
            public Column PlaceFlip { get; set; }
            public Column PlaceStyle { get; set; }
            public Column PlowCode { get; set; }
            public Column ResizePage { get; set; }
            public Column RouteStyle { get; set; }
            public Column AvoidPageBreaks { get; set; }
            public Column DrawingResizeType { get; set; }

            public PageCellQuery() 
            {
                this.PageLeftMargin = this.Columns.Add(VA.ShapeSheet.SRCConstants.PageLeftMargin, "PageLeftMargin");
                this.CenterX = this.Columns.Add(VA.ShapeSheet.SRCConstants.CenterX, "CenterX");
                this.CenterY = this.Columns.Add(VA.ShapeSheet.SRCConstants.CenterY, "CenterY");
                this.OnPage = this.Columns.Add(VA.ShapeSheet.SRCConstants.OnPage, "OnPage");
                this.PageBottomMargin = this.Columns.Add(VA.ShapeSheet.SRCConstants.PageBottomMargin, "PageBottomMargin");
                this.PageRightMargin = this.Columns.Add(VA.ShapeSheet.SRCConstants.PageRightMargin, "PageRightMargin");
                this.PagesX = this.Columns.Add(VA.ShapeSheet.SRCConstants.PagesX, "PagesX");
                this.PagesY = this.Columns.Add(VA.ShapeSheet.SRCConstants.PagesY, "PagesY");
                this.PageTopMargin = this.Columns.Add(VA.ShapeSheet.SRCConstants.PageTopMargin, "PageTopMargin");
                this.PaperKind = this.Columns.Add(VA.ShapeSheet.SRCConstants.PaperKind, "PaperKind");
                this.PrintGrid = this.Columns.Add(VA.ShapeSheet.SRCConstants.PrintGrid, "PrintGrid");
                this.PrintPageOrientation = this.Columns.Add(VA.ShapeSheet.SRCConstants.PrintPageOrientation, "PrintPageOrientation");
                this.ScaleX = this.Columns.Add(VA.ShapeSheet.SRCConstants.ScaleX, "ScaleX");
                this.ScaleY = this.Columns.Add(VA.ShapeSheet.SRCConstants.ScaleY, "ScaleY");
                this.PaperSource = this.Columns.Add(VA.ShapeSheet.SRCConstants.PaperSource, "PaperSource");
                this.DrawingScale = this.Columns.Add(VA.ShapeSheet.SRCConstants.DrawingScale, "DrawingScale");
                this.DrawingScaleType = this.Columns.Add(VA.ShapeSheet.SRCConstants.DrawingScaleType, "DrawingScaleType");
                this.DrawingSizeType = this.Columns.Add(VA.ShapeSheet.SRCConstants.DrawingSizeType, "DrawingSizeType");
                this.InhibitSnap = this.Columns.Add(VA.ShapeSheet.SRCConstants.InhibitSnap, "InhibitSnap");
                this.PageHeight = this.Columns.Add(VA.ShapeSheet.SRCConstants.PageHeight, "PageHeight");
                this.PageScale = this.Columns.Add(VA.ShapeSheet.SRCConstants.PageScale, "PageScale");
                this.PageWidth = this.Columns.Add(VA.ShapeSheet.SRCConstants.PageWidth, "PageWidth");
                this.ShdwObliqueAngle = this.Columns.Add(VA.ShapeSheet.SRCConstants.ShdwObliqueAngle, "ShdwObliqueAngle");
                this.ShdwOffsetX = this.Columns.Add(VA.ShapeSheet.SRCConstants.ShdwOffsetX, "ShdwOffsetX");
                this.ShdwOffsetY = this.Columns.Add(VA.ShapeSheet.SRCConstants.ShdwOffsetY, "ShdwOffsetY");
                this.ShdwScaleFactor = this.Columns.Add(VA.ShapeSheet.SRCConstants.ShdwScaleFactor, "ShdwScaleFactor");
                this.ShdwType = this.Columns.Add(VA.ShapeSheet.SRCConstants.ShdwType, "ShdwType");
                this.UIVisibility = this.Columns.Add(VA.ShapeSheet.SRCConstants.UIVisibility, "UIVisibility");
                this.XGridDensity = this.Columns.Add(VA.ShapeSheet.SRCConstants.XGridDensity, "XGridDensity");
                this.XGridOrigin = this.Columns.Add(VA.ShapeSheet.SRCConstants.XGridOrigin, "XGridOrigin");
                this.XGridSpacing = this.Columns.Add(VA.ShapeSheet.SRCConstants.XGridSpacing, "XGridSpacing");
                this.XRulerDensity = this.Columns.Add(VA.ShapeSheet.SRCConstants.XRulerDensity, "XRulerDensity");
                this.XRulerOrigin = this.Columns.Add(VA.ShapeSheet.SRCConstants.XRulerOrigin, "XRulerOrigin");
                this.YGridDensity = this.Columns.Add(VA.ShapeSheet.SRCConstants.YGridDensity, "YGridDensity");
                this.YGridOrigin = this.Columns.Add(VA.ShapeSheet.SRCConstants.YGridOrigin, "YGridOrigin");
                this.YGridSpacing = this.Columns.Add(VA.ShapeSheet.SRCConstants.YGridSpacing, "YGridSpacing");
                this.YRulerDensity = this.Columns.Add(VA.ShapeSheet.SRCConstants.YRulerDensity, "YRulerDensity");
                this.YRulerOrigin = this.Columns.Add(VA.ShapeSheet.SRCConstants.YRulerOrigin, "YRulerOrigin");
                this.AvenueSizeX = this.Columns.Add(VA.ShapeSheet.SRCConstants.AvenueSizeX, "AvenueSizeX");
                this.AvenueSizeY = this.Columns.Add(VA.ShapeSheet.SRCConstants.AvenueSizeY, "AvenueSizeY");
                this.BlockSizeX = this.Columns.Add(VA.ShapeSheet.SRCConstants.BlockSizeX, "BlockSizeX");
                this.BlockSizeY = this.Columns.Add(VA.ShapeSheet.SRCConstants.BlockSizeY, "BlockSizeY");
                this.CtrlAsInput = this.Columns.Add(VA.ShapeSheet.SRCConstants.CtrlAsInput, "CtrlAsInput");
                this.DynamicsOff = this.Columns.Add(VA.ShapeSheet.SRCConstants.DynamicsOff, "DynamicsOff");
                this.EnableGrid = this.Columns.Add(VA.ShapeSheet.SRCConstants.EnableGrid, "EnableGrid");
                this.LineAdjustFrom = this.Columns.Add(VA.ShapeSheet.SRCConstants.LineAdjustFrom, "LineAdjustFrom");
                this.LineAdjustTo = this.Columns.Add(VA.ShapeSheet.SRCConstants.LineAdjustTo, "LineAdjustTo");
                this.LineJumpCode = this.Columns.Add(VA.ShapeSheet.SRCConstants.LineJumpCode, "LineJumpCode");
                this.LineJumpFactorX = this.Columns.Add(VA.ShapeSheet.SRCConstants.LineJumpFactorX, "LineJumpFactorX");
                this.LineJumpFactorY = this.Columns.Add(VA.ShapeSheet.SRCConstants.LineJumpFactorY, "LineJumpFactorY");
                this.LineJumpStyle = this.Columns.Add(VA.ShapeSheet.SRCConstants.LineJumpStyle, "LineJumpStyle");
                this.LineRouteExt = this.Columns.Add(VA.ShapeSheet.SRCConstants.LineRouteExt, "LineRouteExt");
                this.LineToLineX = this.Columns.Add(VA.ShapeSheet.SRCConstants.LineToLineX, "LineToLineX");
                this.LineToLineY = this.Columns.Add(VA.ShapeSheet.SRCConstants.LineToLineY, "LineToLineY");
                this.LineToNodeX = this.Columns.Add(VA.ShapeSheet.SRCConstants.LineToNodeX, "LineToNodeX");
                this.LineToNodeY = this.Columns.Add(VA.ShapeSheet.SRCConstants.LineToNodeY, "LineToNodeY");
                this.PageLineJumpDirX = this.Columns.Add(VA.ShapeSheet.SRCConstants.PageLineJumpDirX, "PageLineJumpDirX");
                this.PageLineJumpDirY = this.Columns.Add(VA.ShapeSheet.SRCConstants.PageLineJumpDirY, "PageLineJumpDirY");
                this.PageShapeSplit = this.Columns.Add(VA.ShapeSheet.SRCConstants.PageShapeSplit, "PageShapeSplit");
                this.PlaceDepth = this.Columns.Add(VA.ShapeSheet.SRCConstants.PlaceDepth, "PlaceDepth");
                this.PlaceFlip = this.Columns.Add(VA.ShapeSheet.SRCConstants.PlaceFlip, "PlaceFlip");
                this.PlaceStyle = this.Columns.Add(VA.ShapeSheet.SRCConstants.PlaceStyle, "PlaceStyle");
                this.PlowCode = this.Columns.Add(VA.ShapeSheet.SRCConstants.PlowCode, "PlowCode");
                this.ResizePage = this.Columns.Add(VA.ShapeSheet.SRCConstants.ResizePage, "ResizePage");
                this.RouteStyle = this.Columns.Add(VA.ShapeSheet.SRCConstants.RouteStyle, "RouteStyle");
                this.AvoidPageBreaks = this.Columns.Add(VA.ShapeSheet.SRCConstants.AvoidPageBreaks, "AvoidPageBreaks");
                this.DrawingResizeType = this.Columns.Add(VA.ShapeSheet.SRCConstants.DrawingResizeType, "DrawingResizeType");
            }


            public PageCells GetCells(QueryResult<VA.ShapeSheet.CellData<double>> row)
            {
                var cells = new PageCells();
                cells.PageLeftMargin = row[PageLeftMargin.Ordinal];
                cells.CenterX = row[CenterX.Ordinal];
                cells.CenterY = row[CenterY.Ordinal];
                cells.OnPage = row[OnPage.Ordinal].ToInt();
                cells.PageBottomMargin = row[PageBottomMargin.Ordinal];
                cells.PageRightMargin = row[PageRightMargin.Ordinal];
                cells.PagesX = row[PagesX.Ordinal];
                cells.PagesY = row[PagesY.Ordinal];
                cells.PageTopMargin = row[PageTopMargin.Ordinal];
                cells.PaperKind = row[PaperKind.Ordinal].ToInt();
                cells.PrintGrid = row[PrintGrid.Ordinal].ToInt();
                cells.PrintPageOrientation = row[PrintPageOrientation.Ordinal].ToInt();
                cells.ScaleX = row[ScaleX.Ordinal];
                cells.ScaleY = row[ScaleY.Ordinal];
                cells.PaperSource = row[PaperSource.Ordinal].ToInt();
                cells.DrawingScale = row[DrawingScale.Ordinal];
                cells.DrawingScaleType = row[DrawingScaleType.Ordinal].ToInt();
                cells.DrawingSizeType = row[DrawingSizeType.Ordinal].ToInt();
                cells.InhibitSnap = row[InhibitSnap.Ordinal].ToInt();
                cells.PageHeight = row[PageHeight.Ordinal];
                cells.PageScale = row[PageScale.Ordinal];
                cells.PageWidth = row[PageWidth.Ordinal];
                cells.ShdwObliqueAngle = row[ShdwObliqueAngle.Ordinal];
                cells.ShdwOffsetX = row[ShdwOffsetX.Ordinal];
                cells.ShdwOffsetY = row[ShdwOffsetY.Ordinal];
                cells.ShdwScaleFactor = row[ShdwScaleFactor.Ordinal];
                cells.ShdwType = row[ShdwType.Ordinal].ToInt();
                cells.UIVisibility = row[UIVisibility.Ordinal];
                cells.XGridDensity = row[XGridDensity.Ordinal];
                cells.XGridOrigin = row[XGridOrigin.Ordinal];
                cells.XGridSpacing = row[XGridSpacing.Ordinal];
                cells.XRulerDensity = row[XRulerDensity.Ordinal];
                cells.XRulerOrigin = row[XRulerOrigin.Ordinal];
                cells.YGridDensity = row[YGridDensity.Ordinal];
                cells.YGridOrigin = row[YGridOrigin.Ordinal];
                cells.YGridSpacing = row[YGridSpacing.Ordinal];
                cells.YRulerDensity = row[YRulerDensity.Ordinal];
                cells.YRulerOrigin = row[YRulerOrigin.Ordinal];
                cells.AvenueSizeX = row[AvenueSizeX.Ordinal];
                cells.AvenueSizeY = row[AvenueSizeY.Ordinal];
                cells.BlockSizeX = row[BlockSizeX.Ordinal];
                cells.BlockSizeY = row[BlockSizeY.Ordinal];
                cells.CtrlAsInput = row[CtrlAsInput.Ordinal].ToInt();
                cells.DynamicsOff = row[DynamicsOff.Ordinal].ToInt();
                cells.EnableGrid = row[EnableGrid.Ordinal].ToInt();
                cells.LineAdjustFrom = row[LineAdjustFrom.Ordinal].ToInt();
                cells.LineAdjustTo = row[LineAdjustTo.Ordinal];
                cells.LineJumpCode = row[LineJumpCode.Ordinal];
                cells.LineJumpFactorX = row[LineJumpFactorX.Ordinal];
                cells.LineJumpFactorY = row[LineJumpFactorY.Ordinal];
                cells.LineJumpStyle = row[LineJumpStyle.Ordinal].ToInt();
                cells.LineRouteExt = row[LineRouteExt.Ordinal];
                cells.LineToLineX = row[LineToLineX.Ordinal];
                cells.LineToLineY = row[LineToLineY.Ordinal];
                cells.LineToNodeX = row[LineToNodeX.Ordinal];
                cells.LineToNodeY = row[LineToNodeY.Ordinal];
                cells.PageLineJumpDirX = row[PageLineJumpDirX.Ordinal];
                cells.PageLineJumpDirY = row[PageLineJumpDirY.Ordinal];
                cells.PageShapeSplit = row[PageShapeSplit.Ordinal].ToInt();
                cells.PlaceDepth = row[PlaceDepth.Ordinal].ToInt();
                cells.PlaceFlip = row[PlaceFlip.Ordinal].ToInt();
                cells.PlaceStyle = row[PlaceStyle.Ordinal].ToInt();
                cells.PlowCode = row[PlowCode.Ordinal].ToInt();
                cells.ResizePage = row[ResizePage.Ordinal].ToInt();
                cells.RouteStyle = row[RouteStyle.Ordinal].ToInt();
                cells.AvoidPageBreaks = row[AvoidPageBreaks.Ordinal].ToInt();
                cells.DrawingResizeType = row[DrawingResizeType.Ordinal].ToInt();
                return cells;
            }

        }

    }
}