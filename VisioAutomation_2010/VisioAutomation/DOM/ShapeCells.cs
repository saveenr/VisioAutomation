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
        public ShapeSheet.FormulaLiteral CharDoubleStrikeThrough { get; set; }
        public ShapeSheet.FormulaLiteral CharDoubleUnderline { get; set; }
        public ShapeSheet.FormulaLiteral CharFont { get; set; }
        public ShapeSheet.FormulaLiteral CharFontScale { get; set; }
        public ShapeSheet.FormulaLiteral CharLangID { get; set; }
        public ShapeSheet.FormulaLiteral CharLetterspace { get; set; }
        public ShapeSheet.FormulaLiteral CharLocale { get; set; }
        public ShapeSheet.FormulaLiteral CharLocalizeFont { get; set; }
        public ShapeSheet.FormulaLiteral CharOverline { get; set; }
        public ShapeSheet.FormulaLiteral CharPerpendicular { get; set; }
        public ShapeSheet.FormulaLiteral CharPos { get; set; }
        public ShapeSheet.FormulaLiteral CharRTLText { get; set; }
        public ShapeSheet.FormulaLiteral CharSize { get; set; }
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
        public ShapeSheet.FormulaLiteral ParaBullet { get; set; }
        public ShapeSheet.FormulaLiteral ParaBulletFont { get; set; }
        public ShapeSheet.FormulaLiteral ParaBulletFontSize { get; set; }
        public ShapeSheet.FormulaLiteral ParaBulletString { get; set; }
        public ShapeSheet.FormulaLiteral ParaFlags { get; set; }
        public ShapeSheet.FormulaLiteral ParaHorizontalAlign { get; set; }
        public ShapeSheet.FormulaLiteral ParaIndentFirst { get; set; }
        public ShapeSheet.FormulaLiteral ParaIndentLeft { get; set; }
        public ShapeSheet.FormulaLiteral ParaIndentRight { get; set; }
        public ShapeSheet.FormulaLiteral ParaLocBulletFont { get; set; }
        public ShapeSheet.FormulaLiteral ParaSpacingAfter { get; set; }
        public ShapeSheet.FormulaLiteral ParaSpacingBefore { get; set; }
        public ShapeSheet.FormulaLiteral ParaSpacingLine { get; set; }
        public ShapeSheet.FormulaLiteral ParaTextPosAfterBullet { get; set; }

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


        public void Apply(ShapeSheet.Update update, short id)
        {
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.Width, this.Width);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.Height, this.Height);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.PinX, this.PinX);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.PinY, this.PinY);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.LocPinX, this.LocPinX);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.LocPinY, this.LocPinY);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.Angle, this.Angle);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.BeginArrow, this.BeginArrow);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.BeginArrowSize, this.BeginArrowSize);

            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.FillBkgnd, this.FillBkgnd);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.FillBkgndTrans, this.FillBkgndTrans);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.FillForegnd, this.FillForegnd);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.FillForegndTrans, this.FillForegndTrans);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.FillPattern, this.FillPattern);

            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.ShapeShdwObliqueAngle, this.ShapeShdwObliqueAngle);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.ShapeShdwOffsetX, this.ShapeShdwOffsetX);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.ShapeShdwOffsetY, this.ShapeShdwOffsetY);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.ShapeShdwScaleFactor, this.ShapeShdwScaleFactor);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.ShapeShdwType, this.ShapeShdwType);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.ShdwBkgnd, this.ShdwBkgnd);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.ShdwBkgndTrans, this.ShdwBkgndTrans);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.ShdwForegnd, this.ShdwForegnd);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.ShdwForegndTrans, this.ShdwForegndTrans);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.ShdwPattern, this.ShdwPattern);

            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.CharCase, this.CharCase);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.CharFont, this.CharFont);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.CharColor, this.CharColor);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.CharSize, this.CharSize);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.CharLetterspace, this.CharLetterspace);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.CharStyle, this.CharStyle);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.CharColorTrans, this.CharTransparency);

            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.EndArrow, this.EndArrow);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.EndArrowSize, this.EndArrowSize);

            // Line
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.LineColor, this.LineColor);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.LineColorTrans, this.LineColorTrans);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.LinePattern, this.LinePattern);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.LineWeight, this.LineWeight);

            // Text
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.BottomMargin, this.BottomMargin);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.DefaultTabStop, this.DefaultTabstop);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.LeftMargin, this.LeftMargin);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.RightMargin, this.RightMargin);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.TextBkgnd, this.TextBkgnd);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.TextBkgndTrans, this.TextBkgndTrans);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.TextDirection, this.TextDirection);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.TopMargin, this.TopMargin);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.VerticalAlign, this.VerticalAlign);

            // Paragraph

            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.Para_Bullet, this.ParaBullet);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.Para_BulletFont, this.ParaBulletFont);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.Para_BulletFontSize, this.ParaBulletFontSize);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.Para_BulletStr, this.ParaBulletString);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.Para_Flags, this.ParaFlags);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.Para_HorzAlign, this.ParaHorizontalAlign);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.Para_IndFirst, this.ParaIndentFirst);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.Para_IndLeft, this.ParaIndentLeft);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.Para_IndRight, this.ParaIndentRight);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.Para_LocalizeBulletFont, this.ParaLocBulletFont);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.Para_SpAfter, this.ParaSpacingAfter);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.Para_SpBefore, this.ParaSpacingBefore);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.Para_SpLine, this.ParaSpacingLine);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.Para_TextPosAfterBullet, this.ParaTextPosAfterBullet);

            // TextXForm
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.TxtAngle, this.TxtAngle);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.TxtHeight, this.TxtHeight);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.TxtLocPinX, this.TxtLocPinX);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.TxtLocPinY, this.TxtLocPinY);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.TxtPinX, this.TxtPinX);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.TxtPinY, this.TxtPinY);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.TxtWidth, this.TxtWidth);

            // ShapeLayout
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.ConFixedCode, this.ConFixedCode);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.ConLineJumpCode, this.ConLineJumpCode);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.ConLineJumpDirX, this.ConLineJumpDirX);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.ConLineJumpDirY, this.ConLineJumpDirY);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.ConLineJumpStyle, this.ConLineJumpStyle);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.ConLineRouteExt, this.ConLineRouteExt);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.ShapeFixedCode, this.ShapeFixedCode);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.ShapePermeablePlace, this.ShapePermeablePlace);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.ShapePermeableX, this.ShapePermeableX);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.ShapePermeableY, this.ShapePermeableY);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.ShapePlaceFlip, this.ShapePlaceFlip);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.ShapePlaceStyle, this.ShapePlaceStyle);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.ShapePlowCode, this.ShapePlowCode);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.ShapeRouteStyle, this.ShapeRouteStyle);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.ShapeSplit, this.ShapeSplit);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.ShapeSplittable, this.ShapeSplittable);
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
            if (this.CharDoubleStrikeThrough.HasValue) { other.CharDoubleStrikeThrough = this.CharDoubleStrikeThrough; }
            if (this.CharDoubleUnderline.HasValue) { other.CharDoubleUnderline = this.CharDoubleUnderline; }
            if (this.CharFont.HasValue) { other.CharFont = this.CharFont; }
            if (this.CharLangID.HasValue) { other.CharLangID = this.CharLangID; }
            if (this.CharLocale.HasValue) { other.CharLocale = this.CharLocale; }
            if (this.CharLocalizeFont.HasValue) { other.CharLocalizeFont = this.CharLocalizeFont; }
            if (this.CharOverline.HasValue) { other.CharOverline = this.CharOverline; }
            if (this.CharPerpendicular.HasValue) { other.CharPerpendicular = this.CharPerpendicular; }
            if (this.CharPos.HasValue) { other.CharPos = this.CharPos; }
            if (this.CharRTLText.HasValue) { other.CharRTLText = this.CharRTLText; }
            if (this.CharFontScale.HasValue) { other.CharFontScale = this.CharFontScale; }
            if (this.CharSize.HasValue) { other.CharSize = this.CharSize; }
            if (this.CharLetterspace.HasValue) { other.CharLetterspace = this.CharLetterspace; }
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
            if (this.ParaBullet.HasValue) { other.ParaBullet = this.ParaBullet; }
            if (this.ParaBulletFont.HasValue) { other.ParaBulletFont = this.ParaBulletFont; }
            if (this.ParaBulletFontSize.HasValue) { other.ParaBulletFontSize = this.ParaBulletFontSize; }
            if (this.ParaBulletString.HasValue) { other.ParaBulletString = this.ParaBulletString; }
            if (this.ParaFlags.HasValue) { other.ParaFlags = this.ParaFlags; }
            if (this.ParaHorizontalAlign.HasValue) { other.ParaHorizontalAlign = this.ParaHorizontalAlign; }
            if (this.ParaIndentFirst.HasValue) { other.ParaIndentFirst = this.ParaIndentFirst; }
            if (this.ParaIndentLeft.HasValue) { other.ParaIndentLeft = this.ParaIndentLeft; }
            if (this.ParaIndentRight.HasValue) { other.ParaIndentRight = this.ParaIndentRight; }
            if (this.ParaLocBulletFont.HasValue) { other.ParaLocBulletFont = this.ParaLocBulletFont; }
            if (this.ParaSpacingAfter.HasValue) { other.ParaSpacingAfter = this.ParaSpacingAfter; }
            if (this.ParaSpacingBefore.HasValue) { other.ParaSpacingBefore = this.ParaSpacingBefore; }
            if (this.ParaSpacingLine.HasValue) { other.ParaSpacingLine = this.ParaSpacingLine; }
            if (this.ParaTextPosAfterBullet.HasValue) { other.ParaTextPosAfterBullet = this.ParaTextPosAfterBullet; }
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