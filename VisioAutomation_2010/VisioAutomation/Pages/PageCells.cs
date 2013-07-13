using System.Collections.Generic;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.Query;
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

        public override void ApplyFormulas(ApplyFormula func)
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
            func(ShapeSheet.SRCConstants.AvoidPageBreaks, this.AvoidPageBreaks.Formula);
            func(ShapeSheet.SRCConstants.DrawingResizeType, this.DrawingResizeType.Formula);
        }


        public static IList<PageCells> GetCells(IVisio.Page page, IList<int> shapeids)
        {
            var query = get_query();
            return VA.ShapeSheet.CellGroups.CellGroup._GetCells(page, shapeids, query, query.GetCells);
        }

        public static PageCells GetCells(IVisio.Shape shape)
        {
            var query = get_query();
            return VA.ShapeSheet.CellGroups.CellGroup._GetCells(shape, query, query.GetCells);
        }

        private static PageQuery m_query;
        private static PageQuery get_query()
        {
            m_query = m_query ?? new PageQuery();
            return m_query;
        }

        class PageQuery : VA.ShapeSheet.Query.QueryEx
        {
            public int PageLeftMargin { get; set; }
            public int CenterX { get; set; }
            public int CenterY { get; set; }
            public int OnPage { get; set; }
            public int PageBottomMargin { get; set; }
            public int PageRightMargin { get; set; }
            public int PagesX { get; set; }
            public int PagesY { get; set; }
            public int PageTopMargin { get; set; }
            public int PaperKind { get; set; }
            public int PrintGrid { get; set; }
            public int PrintPageOrientation { get; set; }
            public int ScaleX { get; set; }
            public int ScaleY { get; set; }
            public int PaperSource { get; set; }
            public int DrawingScale { get; set; }
            public int DrawingScaleType { get; set; }
            public int DrawingSizeType { get; set; }
            public int InhibitSnap { get; set; }
            public int PageHeight { get; set; }
            public int PageScale { get; set; }
            public int PageWidth { get; set; }
            public int ShdwObliqueAngle { get; set; }
            public int ShdwOffsetX { get; set; }
            public int ShdwOffsetY { get; set; }
            public int ShdwScaleFactor { get; set; }
            public int ShdwType { get; set; }
            public int UIVisibility { get; set; }
            public int XGridDensity { get; set; }
            public int XGridOrigin { get; set; }
            public int XGridSpacing { get; set; }
            public int XRulerDensity { get; set; }
            public int XRulerOrigin { get; set; }
            public int YGridDensity { get; set; }
            public int YGridOrigin { get; set; }
            public int YGridSpacing { get; set; }
            public int YRulerDensity { get; set; }
            public int YRulerOrigin { get; set; }
            public int AvenueSizeX { get; set; }
            public int AvenueSizeY { get; set; }
            public int BlockSizeX { get; set; }
            public int BlockSizeY { get; set; }
            public int CtrlAsInput { get; set; }
            public int DynamicsOff { get; set; }
            public int EnableGrid { get; set; }
            public int LineAdjustFrom { get; set; }
            public int LineAdjustTo { get; set; }
            public int LineJumpCode { get; set; }
            public int LineJumpFactorX { get; set; }
            public int LineJumpFactorY { get; set; }
            public int LineJumpStyle { get; set; }
            public int LineRouteExt { get; set; }
            public int LineToLineX { get; set; }
            public int LineToLineY { get; set; }
            public int LineToNodeX { get; set; }
            public int LineToNodeY { get; set; }
            public int PageLineJumpDirX { get; set; }
            public int PageLineJumpDirY { get; set; }
            public int PageShapeSplit { get; set; }
            public int PlaceDepth { get; set; }
            public int PlaceFlip { get; set; }
            public int PlaceStyle { get; set; }
            public int PlowCode { get; set; }
            public int ResizePage { get; set; }
            public int RouteStyle { get; set; }
            public int AvoidPageBreaks { get; set; }
            public int DrawingResizeType { get; set; }

            public PageQuery() 
            {
                this.PageLeftMargin = this.AddColumn2(VA.ShapeSheet.SRCConstants.PageLeftMargin, "PageLeftMargin");
                this.CenterX = this.AddColumn2(VA.ShapeSheet.SRCConstants.CenterX, "CenterX");
                this.CenterY = this.AddColumn2(VA.ShapeSheet.SRCConstants.CenterY, "CenterY");
                this.OnPage = this.AddColumn2(VA.ShapeSheet.SRCConstants.OnPage, "OnPage");
                this.PageBottomMargin = this.AddColumn2(VA.ShapeSheet.SRCConstants.PageBottomMargin, "PageBottomMargin");
                this.PageRightMargin = this.AddColumn2(VA.ShapeSheet.SRCConstants.PageRightMargin, "PageRightMargin");
                this.PagesX = this.AddColumn2(VA.ShapeSheet.SRCConstants.PagesX, "PagesX");
                this.PagesY = this.AddColumn2(VA.ShapeSheet.SRCConstants.PagesY, "PagesY");
                this.PageTopMargin = this.AddColumn2(VA.ShapeSheet.SRCConstants.PageTopMargin, "PageTopMargin");
                this.PaperKind = this.AddColumn2(VA.ShapeSheet.SRCConstants.PaperKind, "PaperKind");
                this.PrintGrid = this.AddColumn2(VA.ShapeSheet.SRCConstants.PrintGrid, "PrintGrid");
                this.PrintPageOrientation = this.AddColumn2(VA.ShapeSheet.SRCConstants.PrintPageOrientation, "PrintPageOrientation");
                this.ScaleX = this.AddColumn2(VA.ShapeSheet.SRCConstants.ScaleX, "ScaleX");
                this.ScaleY = this.AddColumn2(VA.ShapeSheet.SRCConstants.ScaleY, "ScaleY");
                this.PaperSource = this.AddColumn2(VA.ShapeSheet.SRCConstants.PaperSource, "PaperSource");
                this.DrawingScale = this.AddColumn2(VA.ShapeSheet.SRCConstants.DrawingScale, "DrawingScale");
                this.DrawingScaleType = this.AddColumn2(VA.ShapeSheet.SRCConstants.DrawingScaleType, "DrawingScaleType");
                this.DrawingSizeType = this.AddColumn2(VA.ShapeSheet.SRCConstants.DrawingSizeType, "DrawingSizeType");
                this.InhibitSnap = this.AddColumn2(VA.ShapeSheet.SRCConstants.InhibitSnap, "InhibitSnap");
                this.PageHeight = this.AddColumn2(VA.ShapeSheet.SRCConstants.PageHeight, "PageHeight");
                this.PageScale = this.AddColumn2(VA.ShapeSheet.SRCConstants.PageScale, "PageScale");
                this.PageWidth = this.AddColumn2(VA.ShapeSheet.SRCConstants.PageWidth, "PageWidth");
                this.ShdwObliqueAngle = this.AddColumn2(VA.ShapeSheet.SRCConstants.ShdwObliqueAngle, "ShdwObliqueAngle");
                this.ShdwOffsetX = this.AddColumn2(VA.ShapeSheet.SRCConstants.ShdwOffsetX, "ShdwOffsetX");
                this.ShdwOffsetY = this.AddColumn2(VA.ShapeSheet.SRCConstants.ShdwOffsetY, "ShdwOffsetY");
                this.ShdwScaleFactor = this.AddColumn2(VA.ShapeSheet.SRCConstants.ShdwScaleFactor, "ShdwScaleFactor");
                this.ShdwType = this.AddColumn2(VA.ShapeSheet.SRCConstants.ShdwType, "ShdwType");
                this.UIVisibility = this.AddColumn2(VA.ShapeSheet.SRCConstants.UIVisibility, "UIVisibility");
                this.XGridDensity = this.AddColumn2(VA.ShapeSheet.SRCConstants.XGridDensity, "XGridDensity");
                this.XGridOrigin = this.AddColumn2(VA.ShapeSheet.SRCConstants.XGridOrigin, "XGridOrigin");
                this.XGridSpacing = this.AddColumn2(VA.ShapeSheet.SRCConstants.XGridSpacing, "XGridSpacing");
                this.XRulerDensity = this.AddColumn2(VA.ShapeSheet.SRCConstants.XRulerDensity, "XRulerDensity");
                this.XRulerOrigin = this.AddColumn2(VA.ShapeSheet.SRCConstants.XRulerOrigin, "XRulerOrigin");
                this.YGridDensity = this.AddColumn2(VA.ShapeSheet.SRCConstants.YGridDensity, "YGridDensity");
                this.YGridOrigin = this.AddColumn2(VA.ShapeSheet.SRCConstants.YGridOrigin, "YGridOrigin");
                this.YGridSpacing = this.AddColumn2(VA.ShapeSheet.SRCConstants.YGridSpacing, "YGridSpacing");
                this.YRulerDensity = this.AddColumn2(VA.ShapeSheet.SRCConstants.YRulerDensity, "YRulerDensity");
                this.YRulerOrigin = this.AddColumn2(VA.ShapeSheet.SRCConstants.YRulerOrigin, "YRulerOrigin");
                this.AvenueSizeX = this.AddColumn2(VA.ShapeSheet.SRCConstants.AvenueSizeX, "AvenueSizeX");
                this.AvenueSizeY = this.AddColumn2(VA.ShapeSheet.SRCConstants.AvenueSizeY, "AvenueSizeY");
                this.BlockSizeX = this.AddColumn2(VA.ShapeSheet.SRCConstants.BlockSizeX, "BlockSizeX");
                this.BlockSizeY = this.AddColumn2(VA.ShapeSheet.SRCConstants.BlockSizeY, "BlockSizeY");
                this.CtrlAsInput = this.AddColumn2(VA.ShapeSheet.SRCConstants.CtrlAsInput, "CtrlAsInput");
                this.DynamicsOff = this.AddColumn2(VA.ShapeSheet.SRCConstants.DynamicsOff, "DynamicsOff");
                this.EnableGrid = this.AddColumn2(VA.ShapeSheet.SRCConstants.EnableGrid, "EnableGrid");
                this.LineAdjustFrom = this.AddColumn2(VA.ShapeSheet.SRCConstants.LineAdjustFrom, "LineAdjustFrom");
                this.LineAdjustTo = this.AddColumn2(VA.ShapeSheet.SRCConstants.LineAdjustTo, "LineAdjustTo");
                this.LineJumpCode = this.AddColumn2(VA.ShapeSheet.SRCConstants.LineJumpCode, "LineJumpCode");
                this.LineJumpFactorX = this.AddColumn2(VA.ShapeSheet.SRCConstants.LineJumpFactorX, "LineJumpFactorX");
                this.LineJumpFactorY = this.AddColumn2(VA.ShapeSheet.SRCConstants.LineJumpFactorY, "LineJumpFactorY");
                this.LineJumpStyle = this.AddColumn2(VA.ShapeSheet.SRCConstants.LineJumpStyle, "LineJumpStyle");
                this.LineRouteExt = this.AddColumn2(VA.ShapeSheet.SRCConstants.LineRouteExt, "LineRouteExt");
                this.LineToLineX = this.AddColumn2(VA.ShapeSheet.SRCConstants.LineToLineX, "LineToLineX");
                this.LineToLineY = this.AddColumn2(VA.ShapeSheet.SRCConstants.LineToLineY, "LineToLineY");
                this.LineToNodeX = this.AddColumn2(VA.ShapeSheet.SRCConstants.LineToNodeX, "LineToNodeX");
                this.LineToNodeY = this.AddColumn2(VA.ShapeSheet.SRCConstants.LineToNodeY, "LineToNodeY");
                this.PageLineJumpDirX = this.AddColumn2(VA.ShapeSheet.SRCConstants.PageLineJumpDirX, "PageLineJumpDirX");
                this.PageLineJumpDirY = this.AddColumn2(VA.ShapeSheet.SRCConstants.PageLineJumpDirY, "PageLineJumpDirY");
                this.PageShapeSplit = this.AddColumn2(VA.ShapeSheet.SRCConstants.PageShapeSplit, "PageShapeSplit");
                this.PlaceDepth = this.AddColumn2(VA.ShapeSheet.SRCConstants.PlaceDepth, "PlaceDepth");
                this.PlaceFlip = this.AddColumn2(VA.ShapeSheet.SRCConstants.PlaceFlip, "PlaceFlip");
                this.PlaceStyle = this.AddColumn2(VA.ShapeSheet.SRCConstants.PlaceStyle, "PlaceStyle");
                this.PlowCode = this.AddColumn2(VA.ShapeSheet.SRCConstants.PlowCode, "PlowCode");
                this.ResizePage = this.AddColumn2(VA.ShapeSheet.SRCConstants.ResizePage, "ResizePage");
                this.RouteStyle = this.AddColumn2(VA.ShapeSheet.SRCConstants.RouteStyle, "RouteStyle");
                this.AvoidPageBreaks = this.AddColumn2(VA.ShapeSheet.SRCConstants.AvoidPageBreaks, "AvoidPageBreaks");
                this.DrawingResizeType = this.AddColumn2(VA.ShapeSheet.SRCConstants.DrawingResizeType, "DrawingResizeType");
            }


            public PageCells GetCells(ExQueryResult<CellData<double>> data_for_shape)
            {
                                var row = data_for_shape.Cells;

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