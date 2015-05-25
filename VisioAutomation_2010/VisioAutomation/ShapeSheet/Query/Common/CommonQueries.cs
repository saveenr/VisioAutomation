using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using VisioAutomation.Extensions;
using VAQUERY = VisioAutomation.ShapeSheet.Query;

namespace VisioAutomation.ShapeSheet.Common
{
    class ShapeFormatCellQuery : VAQUERY.CellQuery
    {
        public VAQUERY.CellColumn FillBkgnd { get; set; }
        public VAQUERY.CellColumn FillBkgndTrans { get; set; }
        public VAQUERY.CellColumn FillForegnd { get; set; }
        public VAQUERY.CellColumn FillForegndTrans { get; set; }
        public VAQUERY.CellColumn FillPattern { get; set; }
        public VAQUERY.CellColumn ShapeShdwObliqueAngle { get; set; }
        public VAQUERY.CellColumn ShapeShdwOffsetX { get; set; }
        public VAQUERY.CellColumn ShapeShdwOffsetY { get; set; }
        public VAQUERY.CellColumn ShapeShdwScaleFactor { get; set; }
        public VAQUERY.CellColumn ShapeShdwType { get; set; }
        public VAQUERY.CellColumn ShdwBkgnd { get; set; }
        public VAQUERY.CellColumn ShdwBkgndTrans { get; set; }
        public VAQUERY.CellColumn ShdwForegnd { get; set; }
        public VAQUERY.CellColumn ShdwForegndTrans { get; set; }
        public VAQUERY.CellColumn ShdwPattern { get; set; }
        public VAQUERY.CellColumn BeginArrow { get; set; }
        public VAQUERY.CellColumn BeginArrowSize { get; set; }
        public VAQUERY.CellColumn EndArrow { get; set; }
        public VAQUERY.CellColumn EndArrowSize { get; set; }
        public VAQUERY.CellColumn LineColor { get; set; }
        public VAQUERY.CellColumn LineCap { get; set; }
        public VAQUERY.CellColumn LineColorTrans { get; set; }
        public VAQUERY.CellColumn LinePattern { get; set; }
        public VAQUERY.CellColumn LineWeight { get; set; }
        public VAQUERY.CellColumn Rounding { get; set; }

        public ShapeFormatCellQuery()
        {
            this.FillBkgnd = this.AddCell(ShapeSheet.SRCConstants.FillBkgnd, "FillBkgnd");
            this.FillBkgndTrans = this.AddCell(ShapeSheet.SRCConstants.FillBkgndTrans, "FillBkgndTrans");
            this.FillForegnd = this.AddCell(ShapeSheet.SRCConstants.FillForegnd, "FillForegnd");
            this.FillForegndTrans = this.AddCell(ShapeSheet.SRCConstants.FillForegndTrans, "FillForegndTrans");
            this.FillPattern = this.AddCell(ShapeSheet.SRCConstants.FillPattern, "FillPattern");
            this.ShapeShdwObliqueAngle = this.AddCell(ShapeSheet.SRCConstants.ShapeShdwObliqueAngle, "ShapeShdwObliqueAngle");
            this.ShapeShdwOffsetX = this.AddCell(ShapeSheet.SRCConstants.ShapeShdwOffsetX, "ShapeShdwOffsetX");
            this.ShapeShdwOffsetY = this.AddCell(ShapeSheet.SRCConstants.ShapeShdwOffsetY, "ShapeShdwOffsetY");
            this.ShapeShdwScaleFactor = this.AddCell(ShapeSheet.SRCConstants.ShapeShdwScaleFactor, "ShapeShdwScaleFactor");
            this.ShapeShdwType = this.AddCell(ShapeSheet.SRCConstants.ShapeShdwType, "ShapeShdwType");
            this.ShdwBkgnd = this.AddCell(ShapeSheet.SRCConstants.ShdwBkgnd, "ShdwBkgnd");
            this.ShdwBkgndTrans = this.AddCell(ShapeSheet.SRCConstants.ShdwBkgndTrans, "ShdwBkgndTrans");
            this.ShdwForegnd = this.AddCell(ShapeSheet.SRCConstants.ShdwForegnd, "ShdwForegnd");
            this.ShdwForegndTrans = this.AddCell(ShapeSheet.SRCConstants.ShdwForegndTrans, "ShdwForegndTrans");
            this.ShdwPattern = this.AddCell(ShapeSheet.SRCConstants.ShdwPattern, "ShdwPattern");

            this.BeginArrow = this.AddCell(ShapeSheet.SRCConstants.BeginArrow, "BeginArrow");
            this.BeginArrowSize = this.AddCell(ShapeSheet.SRCConstants.BeginArrowSize, "BeginArrowSize");
            this.EndArrow = this.AddCell(ShapeSheet.SRCConstants.EndArrow, "EndArrow");
            this.EndArrowSize = this.AddCell(ShapeSheet.SRCConstants.EndArrowSize, "EndArrowSize");
            this.LineColor = this.AddCell(ShapeSheet.SRCConstants.LineColor, "LineColor");
            this.LineCap = this.AddCell(ShapeSheet.SRCConstants.LineCap, "LineCap");
            this.LineColorTrans = this.AddCell(ShapeSheet.SRCConstants.LineColorTrans, "LineColorTrans");
            this.LinePattern = this.AddCell(ShapeSheet.SRCConstants.LinePattern, "LinePattern");
            this.LineWeight = this.AddCell(ShapeSheet.SRCConstants.LineWeight, "LineWeight");
            this.Rounding = this.AddCell(ShapeSheet.SRCConstants.Rounding, "Rounding");

        }

        public VA.Shapes.FormatCells GetCells(IList<ShapeSheet.CellData<double>> row)
        {
            var cells = new VA.Shapes.FormatCells();
            cells.FillBkgnd = row[this.FillBkgnd].ToInt();
            cells.FillBkgndTrans = row[this.FillBkgndTrans];
            cells.FillForegnd = row[this.FillForegnd].ToInt();
            cells.FillForegndTrans = row[this.FillForegndTrans];
            cells.FillPattern = row[this.FillPattern].ToInt();
            cells.ShapeShdwObliqueAngle = row[this.ShapeShdwObliqueAngle];
            cells.ShapeShdwOffsetX = row[this.ShapeShdwOffsetX];
            cells.ShapeShdwOffsetY = row[this.ShapeShdwOffsetY];
            cells.ShapeShdwScaleFactor = row[this.ShapeShdwScaleFactor];
            cells.ShapeShdwType = row[this.ShapeShdwType].ToInt();
            cells.ShdwBkgnd = row[this.ShdwBkgnd].ToInt();
            cells.ShdwBkgndTrans = row[this.ShdwBkgndTrans];
            cells.ShdwForegnd = row[this.ShdwForegnd].ToInt();
            cells.ShdwForegndTrans = row[this.ShdwForegndTrans];
            cells.ShdwPattern = row[this.ShdwPattern].ToInt();
            cells.BeginArrow = row[this.BeginArrow].ToInt();
            cells.BeginArrowSize = row[this.BeginArrowSize];
            cells.EndArrow = row[this.EndArrow].ToInt();
            cells.EndArrowSize = row[this.EndArrowSize];
            cells.LineCap = row[this.LineCap].ToInt();
            cells.LineColor = row[this.LineColor].ToInt();
            cells.LineColorTrans = row[this.LineColorTrans];
            cells.LinePattern = row[this.LinePattern].ToInt();
            cells.LineWeight = row[this.LineWeight];
            cells.Rounding = row[this.Rounding];
            return cells;
        }

    }


    class LockCellQuery : VAQUERY.CellQuery
    {
        public VAQUERY.CellColumn LockAspect { get; set; }
        public VAQUERY.CellColumn LockBegin { get; set; }
        public VAQUERY.CellColumn LockCalcWH { get; set; }
        public VAQUERY.CellColumn LockCrop { get; set; }
        public VAQUERY.CellColumn LockCustProp { get; set; }
        public VAQUERY.CellColumn LockDelete { get; set; }
        public VAQUERY.CellColumn LockEnd { get; set; }
        public VAQUERY.CellColumn LockFormat { get; set; }
        public VAQUERY.CellColumn LockFromGroupFormat { get; set; }
        public VAQUERY.CellColumn LockGroup { get; set; }
        public VAQUERY.CellColumn LockHeight { get; set; }
        public VAQUERY.CellColumn LockMoveX { get; set; }
        public VAQUERY.CellColumn LockMoveY { get; set; }
        public VAQUERY.CellColumn LockRotate { get; set; }
        public VAQUERY.CellColumn LockSelect { get; set; }
        public VAQUERY.CellColumn LockTextEdit { get; set; }
        public VAQUERY.CellColumn LockThemeColors { get; set; }
        public VAQUERY.CellColumn LockThemeEffects { get; set; }
        public VAQUERY.CellColumn LockVtxEdit { get; set; }
        public VAQUERY.CellColumn LockWidth { get; set; }

        public LockCellQuery()
        {
            this.LockAspect = this.AddCell(ShapeSheet.SRCConstants.LockAspect, "LockAspect");
            this.LockBegin = this.AddCell(ShapeSheet.SRCConstants.LockBegin, "LockBegin");
            this.LockCalcWH = this.AddCell(ShapeSheet.SRCConstants.LockCalcWH, "LockCalcWH");
            this.LockCrop = this.AddCell(ShapeSheet.SRCConstants.LockCrop, "LockCrop");
            this.LockCustProp = this.AddCell(ShapeSheet.SRCConstants.LockCustProp, "LockCustProp");
            this.LockDelete = this.AddCell(ShapeSheet.SRCConstants.LockDelete, "LockDelete");
            this.LockEnd = this.AddCell(ShapeSheet.SRCConstants.LockEnd, "LockEnd");
            this.LockFormat = this.AddCell(ShapeSheet.SRCConstants.LockFormat, "LockFormat");
            this.LockFromGroupFormat = this.AddCell(ShapeSheet.SRCConstants.LockFromGroupFormat, "LockFromGroupFormat");
            this.LockGroup = this.AddCell(ShapeSheet.SRCConstants.LockGroup, "LockGroup");
            this.LockHeight = this.AddCell(ShapeSheet.SRCConstants.LockHeight, "LockHeight");
            this.LockMoveX = this.AddCell(ShapeSheet.SRCConstants.LockMoveX, "LockMoveX");
            this.LockMoveY = this.AddCell(ShapeSheet.SRCConstants.LockMoveY, "LockMoveY");
            this.LockRotate = this.AddCell(ShapeSheet.SRCConstants.LockRotate, "LockRotate");
            this.LockSelect = this.AddCell(ShapeSheet.SRCConstants.LockSelect, "LockSelect");
            this.LockTextEdit = this.AddCell(ShapeSheet.SRCConstants.LockTextEdit, "LockTextEdit");
            this.LockThemeColors = this.AddCell(ShapeSheet.SRCConstants.LockThemeColors, "LockThemeColors");
            this.LockThemeEffects = this.AddCell(ShapeSheet.SRCConstants.LockThemeEffects, "LockThemeEffects");
            this.LockVtxEdit = this.AddCell(ShapeSheet.SRCConstants.LockVtxEdit, "LockVtxEdit");
            this.LockWidth = this.AddCell(ShapeSheet.SRCConstants.LockWidth, "LockWidth");

        }

        public VA.Shapes.LockCells GetCells(IList<ShapeSheet.CellData<double>> row)
        {
            var cells = new VA.Shapes.LockCells();
            cells.LockAspect = row[this.LockAspect].ToBool();
            cells.LockBegin = row[this.LockBegin].ToBool();
            cells.LockCalcWH = row[this.LockCalcWH].ToBool();
            cells.LockCrop = row[this.LockCrop].ToBool();
            cells.LockCustProp = row[this.LockCustProp].ToBool();
            cells.LockDelete = row[this.LockDelete].ToBool();
            cells.LockEnd = row[this.LockEnd].ToBool();
            cells.LockFormat = row[this.LockFormat].ToBool();
            cells.LockFromGroupFormat = row[this.LockFromGroupFormat].ToBool();
            cells.LockGroup = row[this.LockGroup].ToBool();
            cells.LockHeight = row[this.LockHeight].ToBool();
            cells.LockMoveX = row[this.LockMoveX].ToBool();
            cells.LockMoveY = row[this.LockMoveY].ToBool();
            cells.LockRotate = row[this.LockRotate].ToBool();
            cells.LockSelect = row[this.LockSelect].ToBool();
            cells.LockTextEdit = row[this.LockTextEdit].ToBool();
            cells.LockThemeColors = row[this.LockThemeColors].ToBool();
            cells.LockThemeEffects = row[this.LockThemeEffects].ToBool();
            cells.LockVtxEdit = row[this.LockVtxEdit].ToBool();
            cells.LockWidth = row[this.LockWidth].ToBool();
            return cells;
        }
    }


    class XFormCellQuery : VAQUERY.CellQuery
    {
        public VAQUERY.CellColumn Width { get; set; }
        public VAQUERY.CellColumn Height { get; set; }
        public VAQUERY.CellColumn PinX { get; set; }
        public VAQUERY.CellColumn PinY { get; set; }
        public VAQUERY.CellColumn LocPinX { get; set; }
        public VAQUERY.CellColumn LocPinY { get; set; }
        public VAQUERY.CellColumn Angle { get; set; }

        public XFormCellQuery()
        {
            this.PinX = this.AddCell(ShapeSheet.SRCConstants.PinX, "PinX");
            this.PinY = this.AddCell(ShapeSheet.SRCConstants.PinY, "PinY");
            this.LocPinX = this.AddCell(ShapeSheet.SRCConstants.LocPinX, "LocPinX");
            this.LocPinY = this.AddCell(ShapeSheet.SRCConstants.LocPinY, "LocPinY");
            this.Width = this.AddCell(ShapeSheet.SRCConstants.Width, "Width");
            this.Height = this.AddCell(ShapeSheet.SRCConstants.Height, "Height");
            this.Angle = this.AddCell(ShapeSheet.SRCConstants.Angle, "Angle");
        }

        public VA.Shapes.XFormCells GetCells(IList<ShapeSheet.CellData<double>> row)
        {
            var cells = new VA.Shapes.XFormCells
            {
                PinX = row[this.PinX],
                PinY = row[this.PinY],
                LocPinX = row[this.LocPinX],
                LocPinY = row[this.LocPinY],
                Width = row[this.Width],
                Height = row[this.Height],
                Angle = row[this.Angle]
            };
            return cells;
        }
    }


    class PageCellQuery : VAQUERY.CellQuery
    {
        public VAQUERY.CellColumn PageLeftMargin { get; set; }
        public VAQUERY.CellColumn CenterX { get; set; }
        public VAQUERY.CellColumn CenterY { get; set; }
        public VAQUERY.CellColumn OnPage { get; set; }
        public VAQUERY.CellColumn PageBottomMargin { get; set; }
        public VAQUERY.CellColumn PageRightMargin { get; set; }
        public VAQUERY.CellColumn PagesX { get; set; }
        public VAQUERY.CellColumn PagesY { get; set; }
        public VAQUERY.CellColumn PageTopMargin { get; set; }
        public VAQUERY.CellColumn PaperKind { get; set; }
        public VAQUERY.CellColumn PrintGrid { get; set; }
        public VAQUERY.CellColumn PrintPageOrientation { get; set; }
        public VAQUERY.CellColumn ScaleX { get; set; }
        public VAQUERY.CellColumn ScaleY { get; set; }
        public VAQUERY.CellColumn PaperSource { get; set; }
        public VAQUERY.CellColumn DrawingScale { get; set; }
        public VAQUERY.CellColumn DrawingScaleType { get; set; }
        public VAQUERY.CellColumn DrawingSizeType { get; set; }
        public VAQUERY.CellColumn InhibitSnap { get; set; }
        public VAQUERY.CellColumn PageHeight { get; set; }
        public VAQUERY.CellColumn PageScale { get; set; }
        public VAQUERY.CellColumn PageWidth { get; set; }
        public VAQUERY.CellColumn ShdwObliqueAngle { get; set; }
        public VAQUERY.CellColumn ShdwOffsetX { get; set; }
        public VAQUERY.CellColumn ShdwOffsetY { get; set; }
        public VAQUERY.CellColumn ShdwScaleFactor { get; set; }
        public VAQUERY.CellColumn ShdwType { get; set; }
        public VAQUERY.CellColumn UIVisibility { get; set; }
        public VAQUERY.CellColumn XGridDensity { get; set; }
        public VAQUERY.CellColumn XGridOrigin { get; set; }
        public VAQUERY.CellColumn XGridSpacing { get; set; }
        public VAQUERY.CellColumn XRulerDensity { get; set; }
        public VAQUERY.CellColumn XRulerOrigin { get; set; }
        public VAQUERY.CellColumn YGridDensity { get; set; }
        public VAQUERY.CellColumn YGridOrigin { get; set; }
        public VAQUERY.CellColumn YGridSpacing { get; set; }
        public VAQUERY.CellColumn YRulerDensity { get; set; }
        public VAQUERY.CellColumn YRulerOrigin { get; set; }
        public VAQUERY.CellColumn AvenueSizeX { get; set; }
        public VAQUERY.CellColumn AvenueSizeY { get; set; }
        public VAQUERY.CellColumn BlockSizeX { get; set; }
        public VAQUERY.CellColumn BlockSizeY { get; set; }
        public VAQUERY.CellColumn CtrlAsInput { get; set; }
        public VAQUERY.CellColumn DynamicsOff { get; set; }
        public VAQUERY.CellColumn EnableGrid { get; set; }
        public VAQUERY.CellColumn LineAdjustFrom { get; set; }
        public VAQUERY.CellColumn LineAdjustTo { get; set; }
        public VAQUERY.CellColumn LineJumpCode { get; set; }
        public VAQUERY.CellColumn LineJumpFactorX { get; set; }
        public VAQUERY.CellColumn LineJumpFactorY { get; set; }
        public VAQUERY.CellColumn LineJumpStyle { get; set; }
        public VAQUERY.CellColumn LineRouteExt { get; set; }
        public VAQUERY.CellColumn LineToLineX { get; set; }
        public VAQUERY.CellColumn LineToLineY { get; set; }
        public VAQUERY.CellColumn LineToNodeX { get; set; }
        public VAQUERY.CellColumn LineToNodeY { get; set; }
        public VAQUERY.CellColumn PageLineJumpDirX { get; set; }
        public VAQUERY.CellColumn PageLineJumpDirY { get; set; }
        public VAQUERY.CellColumn PageShapeSplit { get; set; }
        public VAQUERY.CellColumn PlaceDepth { get; set; }
        public VAQUERY.CellColumn PlaceFlip { get; set; }
        public VAQUERY.CellColumn PlaceStyle { get; set; }
        public VAQUERY.CellColumn PlowCode { get; set; }
        public VAQUERY.CellColumn ResizePage { get; set; }
        public VAQUERY.CellColumn RouteStyle { get; set; }
        public VAQUERY.CellColumn AvoidPageBreaks { get; set; }
        public VAQUERY.CellColumn DrawingResizeType { get; set; }

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


        public VA.Pages.PageCells GetCells(IList<ShapeSheet.CellData<double>> row)
        {

            var cells = new VA.Pages.PageCells();
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


    class UserDefinedCellQuery : VAQUERY.CellQuery
    {
        public VAQUERY.CellColumn Value { get; set; }
        public VAQUERY.CellColumn Prompt { get; set; }

        public UserDefinedCellQuery()
        {
            var sec = this.AddSection(IVisio.VisSectionIndices.visSectionUser);
            this.Value = sec.AddCell(ShapeSheet.SRCConstants.User_Value, "User");
            this.Prompt = sec.AddCell(ShapeSheet.SRCConstants.User_Prompt, "Prompt");
        }

        public VA.Shapes.UserDefinedCells.UserDefinedCell GetCells(IList<ShapeSheet.CellData<string>> row)
        {
            var cells = new VA.Shapes.UserDefinedCells.UserDefinedCell();
            cells.Value = row[this.Value];
            cells.Prompt = row[this.Prompt];
            return cells;
        }
    }

    class CharacterFormatCellQuery : VAQUERY.CellQuery
    {
        public VAQUERY.CellColumn Font { get; set; }
        public VAQUERY.CellColumn Style { get; set; }
        public VAQUERY.CellColumn Color { get; set; }
        public VAQUERY.CellColumn Size { get; set; }
        public VAQUERY.CellColumn Trans { get; set; }
        public VAQUERY.CellColumn AsianFont { get; set; }
        public VAQUERY.CellColumn Case { get; set; }
        public VAQUERY.CellColumn ComplexScriptFont { get; set; }
        public VAQUERY.CellColumn ComplexScriptSize { get; set; }
        public VAQUERY.CellColumn DoubleStrikethrough { get; set; }
        public VAQUERY.CellColumn DoubleUnderline { get; set; }
        public VAQUERY.CellColumn LangID { get; set; }
        public VAQUERY.CellColumn Locale { get; set; }
        public VAQUERY.CellColumn LocalizeFont { get; set; }
        public VAQUERY.CellColumn Overline { get; set; }
        public VAQUERY.CellColumn Perpendicular { get; set; }
        public VAQUERY.CellColumn Pos { get; set; }
        public VAQUERY.CellColumn RTLText { get; set; }
        public VAQUERY.CellColumn FontScale { get; set; }
        public VAQUERY.CellColumn Letterspace { get; set; }
        public VAQUERY.CellColumn Strikethru { get; set; }
        public VAQUERY.CellColumn UseVertical { get; set; }

        public CharacterFormatCellQuery()
        {
            var sec = this.AddSection(IVisio.VisSectionIndices.visSectionCharacter);
            this.Color = sec.AddCell(ShapeSheet.SRCConstants.CharColor, "CharColor");
            this.Trans = sec.AddCell(ShapeSheet.SRCConstants.CharColorTrans, "CharColorTrans");
            this.Font = sec.AddCell(ShapeSheet.SRCConstants.CharFont, "CharFont");
            this.Size = sec.AddCell(ShapeSheet.SRCConstants.CharSize, "CharSize");
            this.Style = sec.AddCell(ShapeSheet.SRCConstants.CharStyle, "CharStyle");
            this.AsianFont = sec.AddCell(ShapeSheet.SRCConstants.CharAsianFont, "CharAsianFont");
            this.Case = sec.AddCell(ShapeSheet.SRCConstants.CharCase, "CharCase");
            this.ComplexScriptFont = sec.AddCell(ShapeSheet.SRCConstants.CharComplexScriptFont, "CharComplexScriptFont");
            this.ComplexScriptSize = sec.AddCell(ShapeSheet.SRCConstants.CharComplexScriptSize, "CharComplexScriptSize");
            this.DoubleStrikethrough = sec.AddCell(ShapeSheet.SRCConstants.CharDoubleStrikethrough, "CharDoubleStrikethrough");
            this.DoubleUnderline = sec.AddCell(ShapeSheet.SRCConstants.CharDblUnderline, "CharDblUnderline");
            this.LangID = sec.AddCell(ShapeSheet.SRCConstants.CharLangID, "CharLangID");
            this.Locale = sec.AddCell(ShapeSheet.SRCConstants.CharLocale, "CharLocale");
            this.LocalizeFont = sec.AddCell(ShapeSheet.SRCConstants.CharLocalizeFont, "CharLocalizeFont");
            this.Overline = sec.AddCell(ShapeSheet.SRCConstants.CharOverline, "CharOverline");
            this.Perpendicular = sec.AddCell(ShapeSheet.SRCConstants.CharPerpendicular, "CharPerpendicular");
            this.Pos = sec.AddCell(ShapeSheet.SRCConstants.CharPos, "CharPos");
            this.RTLText = sec.AddCell(ShapeSheet.SRCConstants.CharRTLText, "CharRTLText");
            this.FontScale = sec.AddCell(ShapeSheet.SRCConstants.CharFontScale, "CharFontScale");
            this.Letterspace = sec.AddCell(ShapeSheet.SRCConstants.CharLetterspace, "CharLetterspace");
            this.Strikethru = sec.AddCell(ShapeSheet.SRCConstants.CharStrikethru, "CharStrikethru");
            this.UseVertical = sec.AddCell(ShapeSheet.SRCConstants.CharUseVertical, "CharUseVertical");
        }

        public VA.Text.CharacterCells GetCells(IList<ShapeSheet.CellData<double>> row)
        {
            var cells = new VA.Text.CharacterCells();
            cells.Color = row[this.Color].ToInt();
            cells.Transparency = row[this.Trans];
            cells.Font = row[this.Font].ToInt();
            cells.Size = row[this.Size];
            cells.Style = row[this.Style].ToInt();
            cells.AsianFont = row[this.AsianFont].ToInt();
            cells.AsianFont = row[this.AsianFont].ToInt();
            cells.Case = row[this.Case].ToInt();
            cells.ComplexScriptFont = row[this.ComplexScriptFont].ToInt();
            cells.ComplexScriptSize = row[this.ComplexScriptSize];
            cells.DoubleStrikeThrough = row[this.DoubleStrikethrough].ToBool();
            cells.DoubleUnderline = row[this.DoubleUnderline].ToBool();
            cells.FontScale = row[this.FontScale];
            cells.LangID = row[this.LangID].ToInt();
            cells.Letterspace = row[this.Letterspace];
            cells.Locale = row[this.Locale].ToInt();
            cells.LocalizeFont = row[this.LocalizeFont].ToInt();
            cells.Overline = row[this.Overline].ToBool();
            cells.Perpendicular = row[this.Perpendicular].ToBool();
            cells.Pos = row[this.Pos].ToInt();
            cells.RTLText = row[this.RTLText].ToInt();
            cells.Strikethru = row[this.Strikethru].ToBool();
            cells.UseVertical = row[this.UseVertical].ToInt();

            return cells;
        }


    }

    class ConnectionPointCellQuery : VAQUERY.CellQuery
    {
        public VAQUERY.CellColumn DirX { get; set; }
        public VAQUERY.CellColumn DirY { get; set; }
        public VAQUERY.CellColumn Type { get; set; }
        public VAQUERY.CellColumn X { get; set; }
        public VAQUERY.CellColumn Y { get; set; }

        public ConnectionPointCellQuery()
        {
            var sec = this.AddSection(IVisio.VisSectionIndices.visSectionConnectionPts);
            this.DirX = sec.AddCell(ShapeSheet.SRCConstants.Connections_DirX, "Connections_DirX");
            this.DirY = sec.AddCell(ShapeSheet.SRCConstants.Connections_DirY, "Connections_DirY");
            this.Type = sec.AddCell(ShapeSheet.SRCConstants.Connections_Type, "Connections_Type");
            this.X = sec.AddCell(ShapeSheet.SRCConstants.Connections_X, "Connections_X");
            this.Y = sec.AddCell(ShapeSheet.SRCConstants.Connections_Y, "Connections_Y");
        }

        public VA.Shapes.Connections.ConnectionPointCells GetCells(IList<ShapeSheet.CellData<double>> row)
        {
            var cells = new VA.Shapes.Connections.ConnectionPointCells();
            cells.X = row[this.X];
            cells.Y = row[this.Y];
            cells.DirX = row[this.DirX].ToInt();
            cells.DirY = row[this.DirY].ToInt();
            cells.Type = row[this.Type].ToInt();

            return cells;
        }
    }


    class CustomPropertyCellQuery : VAQUERY.CellQuery
    {
        public VAQUERY.CellColumn SortKey { get; set; }
        public VAQUERY.CellColumn Ask { get; set; }
        public VAQUERY.CellColumn Calendar { get; set; }
        public VAQUERY.CellColumn Format { get; set; }
        public VAQUERY.CellColumn Invis { get; set; }
        public VAQUERY.CellColumn Label { get; set; }
        public VAQUERY.CellColumn LangID { get; set; }
        public VAQUERY.CellColumn Prompt { get; set; }
        public VAQUERY.CellColumn Value { get; set; }
        public VAQUERY.CellColumn Type { get; set; }

        public CustomPropertyCellQuery()
        {
            var sec = this.AddSection(IVisio.VisSectionIndices.visSectionProp);

            this.SortKey = sec.AddCell(ShapeSheet.SRCConstants.Prop_SortKey, "Prop_SortKey");
            this.Ask = sec.AddCell(ShapeSheet.SRCConstants.Prop_Ask, "Prop_Ask");
            this.Calendar = sec.AddCell(ShapeSheet.SRCConstants.Prop_Calendar, "Prop_Calendar");
            this.Format = sec.AddCell(ShapeSheet.SRCConstants.Prop_Format, "Prop_Format");
            this.Invis = sec.AddCell(ShapeSheet.SRCConstants.Prop_Invisible, "Prop_Invisible");
            this.Label = sec.AddCell(ShapeSheet.SRCConstants.Prop_Label, "Prop_Label");
            this.LangID = sec.AddCell(ShapeSheet.SRCConstants.Prop_LangID, "Prop_LangID");
            this.Prompt = sec.AddCell(ShapeSheet.SRCConstants.Prop_Prompt, "Prop_Prompt");
            this.Type = sec.AddCell(ShapeSheet.SRCConstants.Prop_Type, "Prop_Type");
            this.Value = sec.AddCell(ShapeSheet.SRCConstants.Prop_Value, "Prop_Value");

        }

        public VA.Shapes.CustomProperties.CustomPropertyCells GetCells(IList<ShapeSheet.CellData<double>> row)
        {
            var cells = new VA.Shapes.CustomProperties.CustomPropertyCells();
            cells.Value = row[this.Value];
            cells.Calendar = row[this.Calendar].ToInt();
            cells.Format = row[this.Format];
            cells.Invisible = row[this.Invis].ToInt();
            cells.Label = row[this.Label];
            cells.LangId = row[this.LangID].ToInt();
            cells.Prompt = row[this.Prompt];
            cells.SortKey = row[this.SortKey].ToInt();
            cells.Type = row[this.Type].ToInt();
            cells.Ask = row[this.Ask].ToBool();
            return cells;
        }
    }


    class ParagraphFormatCellQuery : VAQUERY.CellQuery
    {
        public VAQUERY.CellColumn Bullet { get; set; }
        public VAQUERY.CellColumn BulletFont { get; set; }
        public VAQUERY.CellColumn BulletFontSize { get; set; }
        public VAQUERY.CellColumn BulletString { get; set; } // NOTE: This is never used
        public VAQUERY.CellColumn Flags { get; set; }
        public VAQUERY.CellColumn HorzAlign { get; set; }
        public VAQUERY.CellColumn IndentFirst { get; set; }
        public VAQUERY.CellColumn IndentLeft { get; set; }
        public VAQUERY.CellColumn IndentRight { get; set; }
        public VAQUERY.CellColumn LocalizeBulletFont { get; set; }
        public VAQUERY.CellColumn SpaceAfter { get; set; }
        public VAQUERY.CellColumn SpaceBefore { get; set; }
        public VAQUERY.CellColumn SpaceLine { get; set; }
        public VAQUERY.CellColumn TextPosAfterBullet { get; set; }

        public ParagraphFormatCellQuery()
        {
            var sec = this.AddSection(IVisio.VisSectionIndices.visSectionParagraph);
            this.Bullet = sec.AddCell(ShapeSheet.SRCConstants.Para_Bullet, "Para_Bullet");
            this.BulletFont = sec.AddCell(ShapeSheet.SRCConstants.Para_BulletFont, "Para_BulletFont");
            this.BulletFontSize = sec.AddCell(ShapeSheet.SRCConstants.Para_BulletFontSize, "Para_BulletFontSize");
            this.BulletString = sec.AddCell(ShapeSheet.SRCConstants.Para_BulletStr, "Para_BulletStr");
            this.Flags = sec.AddCell(ShapeSheet.SRCConstants.Para_Flags, "Para_Flags");
            this.HorzAlign = sec.AddCell(ShapeSheet.SRCConstants.Para_HorzAlign, "Para_HorzAlign");
            this.IndentFirst = sec.AddCell(ShapeSheet.SRCConstants.Para_IndFirst, "Para_IndFirst");
            this.IndentLeft = sec.AddCell(ShapeSheet.SRCConstants.Para_IndLeft, "Para_IndLeft");
            this.IndentRight = sec.AddCell(ShapeSheet.SRCConstants.Para_IndRight, "Para_IndRight");
            this.LocalizeBulletFont = sec.AddCell(ShapeSheet.SRCConstants.Para_LocalizeBulletFont, "Para_LocalizeBulletFont");
            this.SpaceAfter = sec.AddCell(ShapeSheet.SRCConstants.Para_SpAfter, "Para_SpAfter");
            this.SpaceBefore = sec.AddCell(ShapeSheet.SRCConstants.Para_SpBefore, "Para_SpBefore");
            this.SpaceLine = sec.AddCell(ShapeSheet.SRCConstants.Para_SpLine, "Para_SpLine");
            this.TextPosAfterBullet = sec.AddCell(ShapeSheet.SRCConstants.Para_TextPosAfterBullet, "Para_TextPosAfterBullet");
        }

        public VA.Text.ParagraphCells GetCells(IList<ShapeSheet.CellData<double>> row)
        {
            var cells = new VA.Text.ParagraphCells();
            cells.IndentFirst = row[this.IndentFirst];
            cells.IndentLeft = row[this.IndentLeft];
            cells.IndentRight = row[this.IndentRight];
            cells.SpacingAfter = row[this.SpaceAfter];
            cells.SpacingBefore = row[this.SpaceBefore];
            cells.SpacingLine = row[this.SpaceLine];
            cells.HorizontalAlign = row[this.HorzAlign].ToInt();
            cells.Bullet = row[this.Bullet].ToInt();
            cells.BulletFont = row[this.BulletFont].ToInt();
            cells.BulletFontSize = row[this.BulletFontSize].ToInt();
            cells.LocBulletFont = row[this.LocalizeBulletFont].ToInt();
            cells.TextPosAfterBullet = row[this.TextPosAfterBullet];
            cells.Flags = row[this.Flags].ToInt();
            cells.BulletString = ""; // TODO: Figure out some way of getting this

            return cells;
        }
    }

    class TextBlockFormatCellQuery : VAQUERY.CellQuery
    {
        public VAQUERY.CellColumn BottomMargin { get; set; }
        public VAQUERY.CellColumn LeftMargin { get; set; }
        public VAQUERY.CellColumn RightMargin { get; set; }
        public VAQUERY.CellColumn TopMargin { get; set; }
        public VAQUERY.CellColumn DefaultTabStop { get; set; }
        public VAQUERY.CellColumn TextBkgnd { get; set; }
        public VAQUERY.CellColumn TextBkgndTrans { get; set; }
        public VAQUERY.CellColumn TextDirection { get; set; }
        public VAQUERY.CellColumn VerticalAlign { get; set; }
        public VAQUERY.CellColumn TxtWidth { get; set; }
        public VAQUERY.CellColumn TxtHeight { get; set; }
        public VAQUERY.CellColumn TxtPinX { get; set; }
        public VAQUERY.CellColumn TxtPinY { get; set; }
        public VAQUERY.CellColumn TxtLocPinX { get; set; }
        public VAQUERY.CellColumn TxtLocPinY { get; set; }
        public VAQUERY.CellColumn TxtAngle { get; set; }

        public TextBlockFormatCellQuery() :
            base()
        {
            this.BottomMargin = this.AddCell(ShapeSheet.SRCConstants.BottomMargin, "BottomMargin");
            this.LeftMargin = this.AddCell(ShapeSheet.SRCConstants.LeftMargin, "LeftMargin");
            this.RightMargin = this.AddCell(ShapeSheet.SRCConstants.RightMargin, "RightMargin");
            this.TopMargin = this.AddCell(ShapeSheet.SRCConstants.TopMargin, "TopMargin");
            this.DefaultTabStop = this.AddCell(ShapeSheet.SRCConstants.DefaultTabStop, "DefaultTabStop");
            this.TextBkgnd = this.AddCell(ShapeSheet.SRCConstants.TextBkgnd, "TextBkgnd");
            this.TextBkgndTrans = this.AddCell(ShapeSheet.SRCConstants.TextBkgndTrans, "TextBkgndTrans");
            this.TextDirection = this.AddCell(ShapeSheet.SRCConstants.TextDirection, "TextDirection");
            this.VerticalAlign = this.AddCell(ShapeSheet.SRCConstants.VerticalAlign, "VerticalAlign");
            this.TxtPinX = this.AddCell(ShapeSheet.SRCConstants.TxtPinX, "TxtPinX");
            this.TxtPinY = this.AddCell(ShapeSheet.SRCConstants.TxtPinY, "TxtPinY");
            this.TxtLocPinX = this.AddCell(ShapeSheet.SRCConstants.TxtLocPinX, "TxtLocPinX");
            this.TxtLocPinY = this.AddCell(ShapeSheet.SRCConstants.TxtLocPinY, "TxtLocPinY");
            this.TxtWidth = this.AddCell(ShapeSheet.SRCConstants.TxtWidth, "TxtWidth");
            this.TxtHeight = this.AddCell(ShapeSheet.SRCConstants.TxtHeight, "TxtHeight");
            this.TxtAngle = this.AddCell(ShapeSheet.SRCConstants.TxtAngle, "TxtAngle");

        }

        public VA.Text.TextCells GetCells(IList<ShapeSheet.CellData<double>> row)
        {
            var cells = new VA.Text.TextCells();
            cells.BottomMargin = row[this.BottomMargin];
            cells.LeftMargin = row[this.LeftMargin];
            cells.RightMargin = row[this.RightMargin];
            cells.TopMargin = row[this.TopMargin];
            cells.DefaultTabStop = row[this.DefaultTabStop];
            cells.TextBkgnd = row[this.TextBkgnd].ToInt();
            cells.TextBkgndTrans = row[this.TextBkgndTrans];
            cells.TextDirection = row[this.TextDirection].ToInt();
            cells.VerticalAlign = row[this.VerticalAlign].ToInt();
            cells.TxtPinX = row[this.TxtPinX];
            cells.TxtPinY = row[this.TxtPinY];
            cells.TxtLocPinX = row[this.TxtLocPinX];
            cells.TxtLocPinY = row[this.TxtLocPinY];
            cells.TxtWidth = row[this.TxtWidth];
            cells.TxtHeight = row[this.TxtHeight];
            cells.TxtAngle = row[this.TxtAngle];
            return cells;
        }
    }


    class ControlCellQuery : VAQUERY.CellQuery
    {
        public VAQUERY.CellColumn CanGlue { get; set; }
        public VAQUERY.CellColumn Tip { get; set; }
        public VAQUERY.CellColumn X { get; set; }
        public VAQUERY.CellColumn Y { get; set; }
        public VAQUERY.CellColumn YBehavior { get; set; }
        public VAQUERY.CellColumn XBehavior { get; set; }
        public VAQUERY.CellColumn XDynamics { get; set; }
        public VAQUERY.CellColumn YDynamics { get; set; }

        public ControlCellQuery()
        {
            var sec = this.AddSection(IVisio.VisSectionIndices.visSectionControls);
            this.CanGlue = sec.AddCell(ShapeSheet.SRCConstants.Controls_CanGlue, "Controls_CanGlue");
            this.Tip = sec.AddCell(ShapeSheet.SRCConstants.Controls_Tip, "Controls_Tip");
            this.X = sec.AddCell(ShapeSheet.SRCConstants.Controls_X, "Controls_X");
            this.Y = sec.AddCell(ShapeSheet.SRCConstants.Controls_Y, "Controls_Y");
            this.YBehavior = sec.AddCell(ShapeSheet.SRCConstants.Controls_YCon, "Controls_YCon");
            this.XBehavior = sec.AddCell(ShapeSheet.SRCConstants.Controls_XCon, "Controls_XCon");
            this.XDynamics = sec.AddCell(ShapeSheet.SRCConstants.Controls_XDyn, "Controls_XDyn");
            this.YDynamics = sec.AddCell(ShapeSheet.SRCConstants.Controls_YDyn, "Controls_YDyn");
        }

        public VA.Shapes.Controls.ControlCells GetCells(IList<ShapeSheet.CellData<double>> row)
        {
            var cells = new VA.Shapes.Controls.ControlCells();
            cells.CanGlue = row[this.CanGlue].ToInt();
            cells.Tip = row[this.Tip].ToInt();
            cells.X = row[this.X];
            cells.Y = row[this.Y];
            cells.YBehavior = row[this.YBehavior].ToInt();
            cells.XBehavior = row[this.XBehavior].ToInt();
            cells.XDynamics = row[this.XDynamics].ToInt();
            cells.YDynamics = row[this.YDynamics].ToInt();
            return cells;
        }
    }


    class ShapeLayoutCellQuery : VAQUERY.CellQuery
    {
        public VAQUERY.CellColumn ConFixedCode { get; set; }
        public VAQUERY.CellColumn ConLineJumpCode { get; set; }
        public VAQUERY.CellColumn ConLineJumpDirX { get; set; }
        public VAQUERY.CellColumn ConLineJumpDirY { get; set; }
        public VAQUERY.CellColumn ConLineJumpStyle { get; set; }
        public VAQUERY.CellColumn ConLineRouteExt { get; set; }
        public VAQUERY.CellColumn ShapeFixedCode { get; set; }
        public VAQUERY.CellColumn ShapePermeablePlace { get; set; }
        public VAQUERY.CellColumn ShapePermeableX { get; set; }
        public VAQUERY.CellColumn ShapePermeableY { get; set; }
        public VAQUERY.CellColumn ShapePlaceFlip { get; set; }
        public VAQUERY.CellColumn ShapePlaceStyle { get; set; }
        public VAQUERY.CellColumn ShapePlowCode { get; set; }
        public VAQUERY.CellColumn ShapeRouteStyle { get; set; }
        public VAQUERY.CellColumn ShapeSplit { get; set; }
        public VAQUERY.CellColumn ShapeSplittable { get; set; }
        public VAQUERY.CellColumn DisplayLevel { get; set; }
        public VAQUERY.CellColumn Relationships { get; set; }

        public ShapeLayoutCellQuery() :
            base()
        {
            this.ConFixedCode = this.AddCell(ShapeSheet.SRCConstants.ConFixedCode, "ConFixedCode");
            this.ConLineJumpCode = this.AddCell(ShapeSheet.SRCConstants.ConLineJumpCode, "ConLineJumpCode");
            this.ConLineJumpDirX = this.AddCell(ShapeSheet.SRCConstants.ConLineJumpDirX, "ConLineJumpDirX");
            this.ConLineJumpDirY = this.AddCell(ShapeSheet.SRCConstants.ConLineJumpDirY, "ConLineJumpDirY");
            this.ConLineJumpStyle = this.AddCell(ShapeSheet.SRCConstants.ConLineJumpStyle, "ConLineJumpStyle");
            this.ConLineRouteExt = this.AddCell(ShapeSheet.SRCConstants.ConLineRouteExt, "ConLineRouteExt");
            this.ShapeFixedCode = this.AddCell(ShapeSheet.SRCConstants.ShapeFixedCode, "ShapeFixedCode");
            this.ShapePermeablePlace = this.AddCell(ShapeSheet.SRCConstants.ShapePermeablePlace, "ShapePermeablePlace");
            this.ShapePermeableX = this.AddCell(ShapeSheet.SRCConstants.ShapePermeableX, "ShapePermeableX");
            this.ShapePermeableY = this.AddCell(ShapeSheet.SRCConstants.ShapePermeableY, "ShapePermeableY");
            this.ShapePlaceFlip = this.AddCell(ShapeSheet.SRCConstants.ShapePlaceFlip, "ShapePlaceFlip");
            this.ShapePlaceStyle = this.AddCell(ShapeSheet.SRCConstants.ShapePlaceStyle, "ShapePlaceStyle");
            this.ShapePlowCode = this.AddCell(ShapeSheet.SRCConstants.ShapePlowCode, "ShapePlowCode");
            this.ShapeRouteStyle = this.AddCell(ShapeSheet.SRCConstants.ShapeRouteStyle, "ShapeRouteStyle");
            this.ShapeSplit = this.AddCell(ShapeSheet.SRCConstants.ShapeSplit, "ShapeSplit");
            this.ShapeSplittable = this.AddCell(ShapeSheet.SRCConstants.ShapeSplittable, "ShapeSplittable");
            this.DisplayLevel = this.AddCell(ShapeSheet.SRCConstants.DisplayLevel, "DisplayLevel");
            this.Relationships = this.AddCell(ShapeSheet.SRCConstants.Relationships, "Relationships");

        }

        public Shapes.Layout.ShapeLayoutCells GetCells(IList<ShapeSheet.CellData<double>> row)
        {
            var cells = new Shapes.Layout.ShapeLayoutCells();
            cells.ConFixedCode = row[this.ConFixedCode].ToInt();
            cells.ConLineJumpCode = row[this.ConLineJumpCode].ToInt();
            cells.ConLineJumpDirX = row[this.ConLineJumpDirX].ToInt();
            cells.ConLineJumpDirY = row[this.ConLineJumpDirY].ToInt();
            cells.ConLineJumpStyle = row[this.ConLineJumpStyle].ToInt();
            cells.ConLineRouteExt = row[this.ConLineRouteExt].ToInt();
            cells.ShapeFixedCode = row[this.ShapeFixedCode].ToInt();
            cells.ShapePermeablePlace = row[this.ShapePermeablePlace].ToInt();
            cells.ShapePermeableX = row[this.ShapePermeableX].ToInt();
            cells.ShapePermeableY = row[this.ShapePermeableY].ToInt();
            cells.ShapePlaceFlip = row[this.ShapePlaceFlip].ToInt();
            cells.ShapePlaceStyle = row[this.ShapePlaceStyle].ToInt();
            cells.ShapePlowCode = row[this.ShapePlowCode].ToInt();
            cells.ShapeRouteStyle = row[this.ShapeRouteStyle].ToInt();
            cells.ShapeSplit = row[this.ShapeSplit].ToInt();
            cells.ShapeSplittable = row[this.ShapeSplittable].ToInt();
            cells.DisplayLevel = row[this.DisplayLevel].ToInt();
            cells.Relationships = row[this.Relationships].ToInt();
            return cells;
        }
    }

}

