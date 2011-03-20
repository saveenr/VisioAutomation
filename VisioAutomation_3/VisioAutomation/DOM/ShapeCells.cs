using System.Collections.Generic;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.DOM
{
    public class ShapeCells
    {
        // XFormOut
        public ShapeSheet.FormulaLiteral Width { get; set; }
        public ShapeSheet.FormulaLiteral Height { get; set; }
        public ShapeSheet.FormulaLiteral Angle { get; set; }
        public ShapeSheet.FormulaLiteral PinX { get; set; }
        public ShapeSheet.FormulaLiteral PinY { get; set; }
        public ShapeSheet.FormulaLiteral LocPinX { get; set; }
        public ShapeSheet.FormulaLiteral LocPinY { get; set; }

        // Fill
        public ShapeSheet.FormulaLiteral FillBkgnd { get; set; }
        public ShapeSheet.FormulaLiteral FillBkgndTrans { get; set; }
        public ShapeSheet.FormulaLiteral FillForegnd { get; set; }
        public ShapeSheet.FormulaLiteral FillForegndTrans { get; set; }
        public ShapeSheet.FormulaLiteral FillPattern { get; set; }
        public ShapeSheet.FormulaLiteral ShapeShdwObliqueAngle { get; set; }
        public ShapeSheet.FormulaLiteral ShapeShdwOffsetX { get; set; }
        public ShapeSheet.FormulaLiteral ShapeShdwOffsetY { get; set; }
        public ShapeSheet.FormulaLiteral ShapeShdwScaleFactor { get; set; }
        public ShapeSheet.FormulaLiteral ShapeShdwType { get; set; }
        public ShapeSheet.FormulaLiteral ShdwBkgnd { get; set; }
        public ShapeSheet.FormulaLiteral ShdwBkgndTrans { get; set; }
        public ShapeSheet.FormulaLiteral ShdwForegnd { get; set; }
        public ShapeSheet.FormulaLiteral ShdwForegndTrans { get; set; }
        public ShapeSheet.FormulaLiteral ShdwPattern { get; set; }

        // Line
        public ShapeSheet.FormulaLiteral BeginArrow { get; set; }
        public ShapeSheet.FormulaLiteral BeginArrowSize { get; set; }
        public ShapeSheet.FormulaLiteral EndArrow { get; set; }
        public ShapeSheet.FormulaLiteral EndArrowSize { get; set; }
        public ShapeSheet.FormulaLiteral LineCap { get; set; }
        public ShapeSheet.FormulaLiteral LineColor { get; set; }
        public ShapeSheet.FormulaLiteral LineColorTrans { get; set; }
        public ShapeSheet.FormulaLiteral LinePattern { get; set; }
        public ShapeSheet.FormulaLiteral LineWeight { get; set; }
        public ShapeSheet.FormulaLiteral Rounding { get; set; }

        // Character
        public ShapeSheet.FormulaLiteral CharAsianFont { get; set; }
        public ShapeSheet.FormulaLiteral CharCase { get; set; }
        public ShapeSheet.FormulaLiteral CharColor { get; set; }
        public ShapeSheet.FormulaLiteral CharComplexScriptFont { get; set; }
        public ShapeSheet.FormulaLiteral CharComplexScriptSize { get; set; }
        public ShapeSheet.FormulaLiteral CharDoubleStrikethrough { get; set; }
        public ShapeSheet.FormulaLiteral CharDoubleULine { get; set; }
        public ShapeSheet.FormulaLiteral CharFont { get; set; }
        public ShapeSheet.FormulaLiteral CharLangID { get; set; }
        public ShapeSheet.FormulaLiteral CharLocale { get; set; }
        public ShapeSheet.FormulaLiteral CharLocalizeFont { get; set; }
        public ShapeSheet.FormulaLiteral CharOverline { get; set; }
        public ShapeSheet.FormulaLiteral CharPerpendicular { get; set; }
        public ShapeSheet.FormulaLiteral CharPos { get; set; }
        public ShapeSheet.FormulaLiteral CharRTLText { get; set; }
        public ShapeSheet.FormulaLiteral CharScale { get; set; }
        public ShapeSheet.FormulaLiteral CharSize { get; set; }
        public ShapeSheet.FormulaLiteral CharSpacing { get; set; }
        public ShapeSheet.FormulaLiteral CharStrikethru { get; set; }
        public ShapeSheet.FormulaLiteral CharStyle { get; set; }
        public ShapeSheet.FormulaLiteral CharTransparency { get; set; }
        public ShapeSheet.FormulaLiteral CharUseVertical { get; set; }

        // Text Block
        public ShapeSheet.FormulaLiteral BottomMargin { get; set; }
        public ShapeSheet.FormulaLiteral DefaultTabstop { get; set; }
        public ShapeSheet.FormulaLiteral LeftMargin { get; set; }
        public ShapeSheet.FormulaLiteral RightMargin { get; set; }
        public ShapeSheet.FormulaLiteral TextBkgnd { get; set; }
        public ShapeSheet.FormulaLiteral TextBkgndTrans { get; set; }
        public ShapeSheet.FormulaLiteral TextDirection { get; set; }
        public ShapeSheet.FormulaLiteral TopMargin { get; set; }
        public ShapeSheet.FormulaLiteral VerticalAlign { get; set; }

        // Paragraph
        public ShapeSheet.FormulaLiteral Bullet { get; set; }
        public ShapeSheet.FormulaLiteral BulletFont { get; set; }
        public ShapeSheet.FormulaLiteral BulletSize { get; set; }
        public ShapeSheet.FormulaLiteral BulletString { get; set; }
        public ShapeSheet.FormulaLiteral Flags { get; set; }
        public ShapeSheet.FormulaLiteral HAlign { get; set; }
        public ShapeSheet.FormulaLiteral IndFirst { get; set; }
        public ShapeSheet.FormulaLiteral IndLeft { get; set; }
        public ShapeSheet.FormulaLiteral IndRight { get; set; }
        public ShapeSheet.FormulaLiteral LocBulletFont { get; set; }
        public ShapeSheet.FormulaLiteral SpAfter { get; set; }
        public ShapeSheet.FormulaLiteral SpBefore { get; set; }
        public ShapeSheet.FormulaLiteral SpLine { get; set; }
        public ShapeSheet.FormulaLiteral TextPosAfterBullet { get; set; }

        //TextXForm
        public ShapeSheet.FormulaLiteral TxtAngle { get; set; }
        public ShapeSheet.FormulaLiteral TxtHeight { get; set; }
        public ShapeSheet.FormulaLiteral TxtLocPinX { get; set; }
        public ShapeSheet.FormulaLiteral TxtLocPinY { get; set; }
        public ShapeSheet.FormulaLiteral TxtPinX { get; set; }
        public ShapeSheet.FormulaLiteral TxtPinY { get; set; }
        public ShapeSheet.FormulaLiteral TxtWidth { get; set; }

        // ShapeLayout
        public ShapeSheet.FormulaLiteral ConFixedCode { get; set; }
        public ShapeSheet.FormulaLiteral ConLineJumpCode { get; set; }
        public ShapeSheet.FormulaLiteral ConLineJumpDirX { get; set; }
        public ShapeSheet.FormulaLiteral ConLineJumpDirY { get; set; }
        public ShapeSheet.FormulaLiteral ConLineJumpStyle { get; set; }
        public ShapeSheet.FormulaLiteral ConLineRouteExt { get; set; }
        public ShapeSheet.FormulaLiteral ShapeFixedCode { get; set; }
        public ShapeSheet.FormulaLiteral ShapePermeablePlace { get; set; }
        public ShapeSheet.FormulaLiteral ShapePermeableX { get; set; }
        public ShapeSheet.FormulaLiteral ShapePermeableY { get; set; }
        public ShapeSheet.FormulaLiteral ShapePlaceFlip { get; set; }
        public ShapeSheet.FormulaLiteral ShapePlaceStyle { get; set; }
        public ShapeSheet.FormulaLiteral ShapePlowCode { get; set; }
        public ShapeSheet.FormulaLiteral ShapeRouteStyle { get; set; }
        public ShapeSheet.FormulaLiteral ShapeSplit { get; set; }
        public ShapeSheet.FormulaLiteral ShapeSplittable { get; set; }


        public void Apply(VA.ShapeSheet.Update.SIDSRCUpdate update, short id)
        {
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.Width, Width);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.Height, Height);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.PinX, PinX);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.PinY, PinY);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.LocPinX, LocPinX);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.LocPinY, LocPinY);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.Angle, Angle);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.BeginArrow, BeginArrow);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.BeginArrowSize, BeginArrowSize);

            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.FillBkgnd, FillBkgnd);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.FillBkgndTrans, FillBkgndTrans);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.FillForegnd, FillForegnd);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.FillForegndTrans, FillForegndTrans);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.FillPattern, FillPattern);

            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.ShapeShdwObliqueAngle, ShapeShdwObliqueAngle);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.ShapeShdwOffsetX, ShapeShdwOffsetX);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.ShapeShdwOffsetY, ShapeShdwOffsetY);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.ShapeShdwScaleFactor, ShapeShdwScaleFactor);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.ShapeShdwType, ShapeShdwType);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.ShdwBkgnd, ShdwBkgnd);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.ShdwBkgndTrans, ShdwBkgndTrans);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.ShdwForegnd, ShdwForegnd);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.ShdwForegndTrans, ShdwForegndTrans);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.ShdwPattern, ShdwPattern);

            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.Char_Case, CharCase);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.Char_Font, CharFont);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.Char_Color, CharColor);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.Char_Size, CharSize);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.Char_Letterspace, CharSpacing);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.Char_Style, CharStyle);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.Char_ColorTrans, CharTransparency);

            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.EndArrow, EndArrow);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.EndArrowSize, EndArrowSize);

            // Line
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.LineColor, LineColor);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.LineColorTrans, LineColorTrans);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.LinePattern, LinePattern);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.LineWeight, LineWeight);

            // Text
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.BottomMargin, BottomMargin);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.DefaultTabstop, DefaultTabstop);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.LeftMargin, LeftMargin);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.RightMargin, RightMargin);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.TextBkgnd, TextBkgnd);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.TextBkgndTrans, TextBkgndTrans);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.TextDirection, TextDirection);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.TopMargin, TopMargin);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.VerticalAlign, VerticalAlign);

            // Paragraph
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.TxtAngle, TxtAngle);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.TxtHeight, TxtHeight);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.TxtLocPinX, TxtLocPinX);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.TxtLocPinY, TxtLocPinY);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.TxtPinX, TxtPinX);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.TxtPinY, TxtPinY);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.TxtWidth, TxtWidth);

            // ShapeLayout
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.ConFixedCode, ConFixedCode);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.ConLineJumpCode, ConLineJumpCode);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.ConLineJumpDirX, ConLineJumpDirX);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.ConLineJumpDirY, ConLineJumpDirY);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.ConLineJumpStyle, ConLineJumpStyle);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.ConLineRouteExt, ConLineRouteExt);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.ShapeFixedCode, ShapeFixedCode);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.ShapePermeablePlace, ShapePermeablePlace);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.ShapePermeableX, ShapePermeableX);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.ShapePermeableY, ShapePermeableY);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.ShapePlaceFlip, ShapePlaceFlip);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.ShapePlaceStyle, ShapePlaceStyle);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.ShapePlowCode, ShapePlowCode);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.ShapeRouteStyle, ShapeRouteStyle);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.ShapeSplit, ShapeSplit);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.ShapeSplittable, ShapeSplittable);

        }

        public ShapeCells ShallowCopy()
        {
            return (ShapeCells) this.MemberwiseClone();
        }
    }

    public class PageCells
    {
        // PageLayout
        public ShapeSheet.FormulaLiteral AvenueSizeX { get; set; }
        public ShapeSheet.FormulaLiteral AvenueSizeY { get; set; }
        public ShapeSheet.FormulaLiteral BlockSizeX { get; set; }
        public ShapeSheet.FormulaLiteral BlockSizeY { get; set; }
        public ShapeSheet.FormulaLiteral CtrlAsInput { get; set; }
        public ShapeSheet.FormulaLiteral DynamicsOff { get; set; }
        public ShapeSheet.FormulaLiteral EnableGrid { get; set; }
        public ShapeSheet.FormulaLiteral LineAdjustFrom { get; set; }
        public ShapeSheet.FormulaLiteral LineAdjustTo { get; set; }
        public ShapeSheet.FormulaLiteral LineJumpCode { get; set; }
        public ShapeSheet.FormulaLiteral LineJumpFactorX { get; set; }
        public ShapeSheet.FormulaLiteral LineJumpFactorY { get; set; }
        public ShapeSheet.FormulaLiteral LineJumpStyle { get; set; }
        public ShapeSheet.FormulaLiteral LineRouteExt { get; set; }
        public ShapeSheet.FormulaLiteral LineToLineX { get; set; }
        public ShapeSheet.FormulaLiteral LineToLineY { get; set; }
        public ShapeSheet.FormulaLiteral LineToNodeX { get; set; }
        public ShapeSheet.FormulaLiteral LineToNodeY { get; set; }
        public ShapeSheet.FormulaLiteral PageLineJumpDirX { get; set; }
        public ShapeSheet.FormulaLiteral PageLineJumpDirY { get; set; }
        public ShapeSheet.FormulaLiteral PageShapeSplit { get; set; }
        public ShapeSheet.FormulaLiteral PlaceDepth { get; set; }
        public ShapeSheet.FormulaLiteral PlaceFlip { get; set; }
        public ShapeSheet.FormulaLiteral PlaceStyle { get; set; }
        public ShapeSheet.FormulaLiteral PlowCode { get; set; }
        public ShapeSheet.FormulaLiteral ResizePage { get; set; }
        public ShapeSheet.FormulaLiteral RouteStyle { get; set; }

        public void Apply(VA.ShapeSheet.Update.SIDSRCUpdate update, short id)
        {
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.AvenueSizeX, AvenueSizeX);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.AvenueSizeY, AvenueSizeY);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.BlockSizeX, BlockSizeX);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.BlockSizeY, BlockSizeY);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.CtrlAsInput, CtrlAsInput);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.DynamicsOff, DynamicsOff);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.EnableGrid, EnableGrid);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.LineAdjustFrom, LineAdjustFrom);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.LineAdjustTo, LineAdjustTo);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.LineJumpCode, LineJumpCode);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.LineJumpFactorX, LineJumpFactorX);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.LineJumpFactorY, LineJumpFactorY);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.LineJumpStyle, LineJumpStyle);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.LineRouteExt, LineRouteExt);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.LineToLineX, LineToLineX);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.LineToLineY, LineToLineY);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.LineToNodeX, LineToNodeX);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.LineToNodeY, LineToNodeY);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.PageLineJumpDirX, PageLineJumpDirX);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.PageLineJumpDirY, PageLineJumpDirY);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.PageShapeSplit, PageShapeSplit);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.PlaceDepth, PlaceDepth);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.PlaceFlip, PlaceFlip);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.PlaceStyle, PlaceStyle);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.PlowCode, PlowCode);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.ResizePage, ResizePage);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.RouteStyle, RouteStyle);
        }

        public PageCells ShallowCopy()
        {
            return (PageCells) this.MemberwiseClone();
        }
    }
}