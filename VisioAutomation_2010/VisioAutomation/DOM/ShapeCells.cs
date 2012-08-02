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
        public ShapeSheet.FormulaLiteral BulletIndex { get; set; }
        public ShapeSheet.FormulaLiteral BulletFont { get; set; }
        public ShapeSheet.FormulaLiteral BulletFontSize { get; set; }
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


        public void Apply(VA.ShapeSheet.Update update, short id)
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
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.DefaultTabStop, DefaultTabstop);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.LeftMargin, LeftMargin);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.RightMargin, RightMargin);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.TextBkgnd, TextBkgnd);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.TextBkgndTrans, TextBkgndTrans);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.TextDirection, TextDirection);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.TopMargin, TopMargin);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.VerticalAlign, VerticalAlign);

            // Paragraph

            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.Para_Bullet, BulletIndex);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.Para_BulletFont, BulletFont);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.Para_BulletFontSize, BulletFontSize);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.Para_BulletStr, BulletString);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.Para_Flags, Flags);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.Para_HorzAlign, HAlign);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.Para_IndFirst, IndFirst);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.Para_IndLeft, IndLeft);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.Para_IndRight, IndRight);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.Para_LocalizeBulletFont, LocBulletFont);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.Para_SpAfter, SpAfter);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.Para_SpBefore, SpBefore);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.Para_SpLine, SpLine);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.Para_TextPosAfterBullet, TextPosAfterBullet);

            // TextXForm
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

        public void ApplyFormulasTo(ShapeCells other)
        {
            if (this.Width.HasValue) { other.Width = this.Width; }
            if (this.Height.HasValue) { other.Height = this.Height; }
            if (this.Angle.HasValue) { other.Angle = this.Angle; }
            if (this.PinX.HasValue) { other.PinX = this.PinX; }
            if (this.PinY.HasValue) { other.PinY = this.PinY; }
            if (this.LocPinX.HasValue) { other.LocPinX = this.LocPinX; }
            if (this.LocPinY.HasValue) { other.LocPinY = this.LocPinY; }
            if (this.FillBkgnd.HasValue) { other.FillBkgnd = this.FillBkgnd; }
            if (this.FillBkgndTrans.HasValue) { other.FillBkgndTrans = this.FillBkgndTrans; }
            if (this.FillForegnd.HasValue) { other.FillForegnd = this.FillForegnd; }
            if (this.FillForegndTrans.HasValue) { other.FillForegndTrans = this.FillForegndTrans; }
            if (this.FillPattern.HasValue) { other.FillPattern = this.FillPattern; }
            if (this.ShapeShdwObliqueAngle.HasValue) { other.ShapeShdwObliqueAngle = this.ShapeShdwObliqueAngle; }
            if (this.ShapeShdwOffsetX.HasValue) { other.ShapeShdwOffsetX = this.ShapeShdwOffsetX; }
            if (this.ShapeShdwOffsetY.HasValue) { other.ShapeShdwOffsetY = this.ShapeShdwOffsetY; }
            if (this.ShapeShdwScaleFactor.HasValue) { other.ShapeShdwScaleFactor = this.ShapeShdwScaleFactor; }
            if (this.ShapeShdwType.HasValue) { other.ShapeShdwType = this.ShapeShdwType; }
            if (this.ShdwBkgnd.HasValue) { other.ShdwBkgnd = this.ShdwBkgnd; }
            if (this.ShdwBkgndTrans.HasValue) { other.ShdwBkgndTrans = this.ShdwBkgndTrans; }
            if (this.ShdwForegnd.HasValue) { other.ShdwForegnd = this.ShdwForegnd; }
            if (this.ShdwForegndTrans.HasValue) { other.ShdwForegndTrans = this.ShdwForegndTrans; }
            if (this.ShdwPattern.HasValue) { other.ShdwPattern = this.ShdwPattern; }
            if (this.BeginArrow.HasValue) { other.BeginArrow = this.BeginArrow; }
            if (this.BeginArrowSize.HasValue) { other.BeginArrowSize = this.BeginArrowSize; }
            if (this.EndArrow.HasValue) { other.EndArrow = this.EndArrow; }
            if (this.EndArrowSize.HasValue) { other.EndArrowSize = this.EndArrowSize; }
            if (this.LineCap.HasValue) { other.LineCap = this.LineCap; }
            if (this.LineColor.HasValue) { other.LineColor = this.LineColor; }
            if (this.LineColorTrans.HasValue) { other.LineColorTrans = this.LineColorTrans; }
            if (this.LinePattern.HasValue) { other.LinePattern = this.LinePattern; }
            if (this.LineWeight.HasValue) { other.LineWeight = this.LineWeight; }
            if (this.Rounding.HasValue) { other.Rounding = this.Rounding; }
            if (this.CharAsianFont.HasValue) { other.CharAsianFont = this.CharAsianFont; }
            if (this.CharCase.HasValue) { other.CharCase = this.CharCase; }
            if (this.CharColor.HasValue) { other.CharColor = this.CharColor; }
            if (this.CharComplexScriptFont.HasValue) { other.CharComplexScriptFont = this.CharComplexScriptFont; }
            if (this.CharComplexScriptSize.HasValue) { other.CharComplexScriptSize = this.CharComplexScriptSize; }
            if (this.CharDoubleStrikethrough.HasValue) { other.CharDoubleStrikethrough = this.CharDoubleStrikethrough; }
            if (this.CharDoubleULine.HasValue) { other.CharDoubleULine = this.CharDoubleULine; }
            if (this.CharFont.HasValue) { other.CharFont = this.CharFont; }
            if (this.CharLangID.HasValue) { other.CharLangID = this.CharLangID; }
            if (this.CharLocale.HasValue) { other.CharLocale = this.CharLocale; }
            if (this.CharLocalizeFont.HasValue) { other.CharLocalizeFont = this.CharLocalizeFont; }
            if (this.CharOverline.HasValue) { other.CharOverline = this.CharOverline; }
            if (this.CharPerpendicular.HasValue) { other.CharPerpendicular = this.CharPerpendicular; }
            if (this.CharPos.HasValue) { other.CharPos = this.CharPos; }
            if (this.CharRTLText.HasValue) { other.CharRTLText = this.CharRTLText; }
            if (this.CharScale.HasValue) { other.CharScale = this.CharScale; }
            if (this.CharSize.HasValue) { other.CharSize = this.CharSize; }
            if (this.CharSpacing.HasValue) { other.CharSpacing = this.CharSpacing; }
            if (this.CharStrikethru.HasValue) { other.CharStrikethru = this.CharStrikethru; }
            if (this.CharStyle.HasValue) { other.CharStyle = this.CharStyle; }
            if (this.CharTransparency.HasValue) { other.CharTransparency = this.CharTransparency; }
            if (this.CharUseVertical.HasValue) { other.CharUseVertical = this.CharUseVertical; }
            if (this.BottomMargin.HasValue) { other.BottomMargin = this.BottomMargin; }
            if (this.DefaultTabstop.HasValue) { other.DefaultTabstop = this.DefaultTabstop; }
            if (this.LeftMargin.HasValue) { other.LeftMargin = this.LeftMargin; }
            if (this.RightMargin.HasValue) { other.RightMargin = this.RightMargin; }
            if (this.TextBkgnd.HasValue) { other.TextBkgnd = this.TextBkgnd; }
            if (this.TextBkgndTrans.HasValue) { other.TextBkgndTrans = this.TextBkgndTrans; }
            if (this.TextDirection.HasValue) { other.TextDirection = this.TextDirection; }
            if (this.TopMargin.HasValue) { other.TopMargin = this.TopMargin; }
            if (this.VerticalAlign.HasValue) { other.VerticalAlign = this.VerticalAlign; }
            if (this.BulletIndex.HasValue) { other.BulletIndex = this.BulletIndex; }
            if (this.BulletFont.HasValue) { other.BulletFont = this.BulletFont; }
            if (this.BulletFontSize.HasValue) { other.BulletFontSize = this.BulletFontSize; }
            if (this.BulletString.HasValue) { other.BulletString = this.BulletString; }
            if (this.Flags.HasValue) { other.Flags = this.Flags; }
            if (this.HAlign.HasValue) { other.HAlign = this.HAlign; }
            if (this.IndFirst.HasValue) { other.IndFirst = this.IndFirst; }
            if (this.IndLeft.HasValue) { other.IndLeft = this.IndLeft; }
            if (this.IndRight.HasValue) { other.IndRight = this.IndRight; }
            if (this.LocBulletFont.HasValue) { other.LocBulletFont = this.LocBulletFont; }
            if (this.SpAfter.HasValue) { other.SpAfter = this.SpAfter; }
            if (this.SpBefore.HasValue) { other.SpBefore = this.SpBefore; }
            if (this.SpLine.HasValue) { other.SpLine = this.SpLine; }
            if (this.TextPosAfterBullet.HasValue) { other.TextPosAfterBullet = this.TextPosAfterBullet; }
            if (this.TxtAngle.HasValue) { other.TxtAngle = this.TxtAngle; }
            if (this.TxtHeight.HasValue) { other.TxtHeight = this.TxtHeight; }
            if (this.TxtLocPinX.HasValue) { other.TxtLocPinX = this.TxtLocPinX; }
            if (this.TxtLocPinY.HasValue) { other.TxtLocPinY = this.TxtLocPinY; }
            if (this.TxtPinX.HasValue) { other.TxtPinX = this.TxtPinX; }
            if (this.TxtPinY.HasValue) { other.TxtPinY = this.TxtPinY; }
            if (this.TxtWidth.HasValue) { other.TxtWidth = this.TxtWidth; }
            if (this.ConFixedCode.HasValue) { other.ConFixedCode = this.ConFixedCode; }
            if (this.ConLineJumpCode.HasValue) { other.ConLineJumpCode = this.ConLineJumpCode; }
            if (this.ConLineJumpDirX.HasValue) { other.ConLineJumpDirX = this.ConLineJumpDirX; }
            if (this.ConLineJumpDirY.HasValue) { other.ConLineJumpDirY = this.ConLineJumpDirY; }
            if (this.ConLineJumpStyle.HasValue) { other.ConLineJumpStyle = this.ConLineJumpStyle; }
            if (this.ConLineRouteExt.HasValue) { other.ConLineRouteExt = this.ConLineRouteExt; }
            if (this.ShapeFixedCode.HasValue) { other.ShapeFixedCode = this.ShapeFixedCode; }
            if (this.ShapePermeablePlace.HasValue) { other.ShapePermeablePlace = this.ShapePermeablePlace; }
            if (this.ShapePermeableX.HasValue) { other.ShapePermeableX = this.ShapePermeableX; }
            if (this.ShapePermeableY.HasValue) { other.ShapePermeableY = this.ShapePermeableY; }
            if (this.ShapePlaceFlip.HasValue) { other.ShapePlaceFlip = this.ShapePlaceFlip; }
            if (this.ShapePlaceStyle.HasValue) { other.ShapePlaceStyle = this.ShapePlaceStyle; }
            if (this.ShapePlowCode.HasValue) { other.ShapePlowCode = this.ShapePlowCode; }
            if (this.ShapeRouteStyle.HasValue) { other.ShapeRouteStyle = this.ShapeRouteStyle; }
            if (this.ShapeSplit.HasValue) { other.ShapeSplit = this.ShapeSplit; }
            if (this.ShapeSplittable.HasValue) { other.ShapeSplittable = this.ShapeSplittable; }
        }
    }
}