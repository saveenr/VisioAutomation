using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Extensions;
using IVisio=Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;

namespace VisioAutomation.Layout
{

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