using VisioAutomation.ShapeSheet.Writers;

namespace VisioAutomation.Models.Dom
{
    public class ShapeCells
    {
        // XFormOut
        public Core.CellValue XFormWidth { get; set; }
        public Core.CellValue XFormHeight { get; set; }
        public Core.CellValue XFormAngle { get; set; }
        public Core.CellValue XFormPinX { get; set; }
        public Core.CellValue XFormPinY { get; set; }
        public Core.CellValue XFormLocPinX { get; set; }
        public Core.CellValue XFormLocPinY { get; set; }

        // Fill
        public Core.CellValue FillBackground { get; set; }
        public Core.CellValue FillBackgroundTransparency { get; set; }
        public Core.CellValue FillForeground { get; set; }
        public Core.CellValue FillForegroundTransparency { get; set; }
        public Core.CellValue FillPattern { get; set; }
        public Core.CellValue FillShadowObliqueAngle { get; set; }
        public Core.CellValue FillShadowOffsetX { get; set; }
        public Core.CellValue FillShadowOffsetY { get; set; }
        public Core.CellValue FillShadowScaleFactor { get; set; }
        public Core.CellValue FillShadowType { get; set; }
        public Core.CellValue FillShadowBackground { get; set; }
        public Core.CellValue FillShadowBackgroundTransparency { get; set; }
        public Core.CellValue FillShadowForeground { get; set; }
        public Core.CellValue FillShadowForegroundTransparency { get; set; }
        public Core.CellValue FillShadowPattern { get; set; }

        // Line
        public Core.CellValue LineBeginArrow { get; set; }
        public Core.CellValue LineBeginArrowSize { get; set; }
        public Core.CellValue LineEndArrow { get; set; }
        public Core.CellValue LineEndArrowSize { get; set; }
        public Core.CellValue LineCap { get; set; }
        public Core.CellValue LineColor { get; set; }
        public Core.CellValue LineColorTransparency { get; set; }
        public Core.CellValue LinePattern { get; set; }
        public Core.CellValue LineWeight { get; set; }
        public Core.CellValue Rounding { get; set; }

        // Character
        public Core.CellValue CharAsianFont { get; set; }
        public Core.CellValue CharCase { get; set; }
        public Core.CellValue CharColor { get; set; }
        public Core.CellValue CharComplexScriptFont { get; set; }
        public Core.CellValue CharComplexScriptSize { get; set; }
        public Core.CellValue CharDoubleStrikeThrough { get; set; }
        public Core.CellValue CharDoubleUnderline { get; set; }
        public Core.CellValue CharFont { get; set; }
        public Core.CellValue CharFontScale { get; set; }
        public Core.CellValue CharLangID { get; set; }
        public Core.CellValue CharLetterspace { get; set; }
        public Core.CellValue CharLocale { get; set; }
        public Core.CellValue CharLocalizeFont { get; set; }
        public Core.CellValue CharOverline { get; set; }
        public Core.CellValue CharPerpendicular { get; set; }
        public Core.CellValue CharPos { get; set; }
        public Core.CellValue CharRTLText { get; set; }
        public Core.CellValue CharSize { get; set; }
        public Core.CellValue CharStrikethru { get; set; }
        public Core.CellValue CharStyle { get; set; }
        public Core.CellValue CharTransparency { get; set; }
        public Core.CellValue CharUseVertical { get; set; }

        // Text Block
        public Core.CellValue TextBlockBottomMargin { get; set; }
        public Core.CellValue TextBlockDefaultTabStop { get; set; }
        public Core.CellValue TextBlockLeftMargin { get; set; }
        public Core.CellValue TextBlockRightMargin { get; set; }
        public Core.CellValue TextBlockBackground { get; set; }
        public Core.CellValue TextBlockBackgroundTransparency { get; set; }
        public Core.CellValue TextBlockDirection { get; set; }
        public Core.CellValue TextBlockTopMargin { get; set; }
        public Core.CellValue TextBlockVerticalAlign { get; set; }

        // Paragraph
        public Core.CellValue ParaBullet { get; set; }
        public Core.CellValue ParaBulletFont { get; set; }
        public Core.CellValue ParaBulletFontSize { get; set; }
        public Core.CellValue ParaBulletString { get; set; }
        public Core.CellValue ParaFlags { get; set; }
        public Core.CellValue ParaHorizontalAlign { get; set; }
        public Core.CellValue ParaIndentFirst { get; set; }
        public Core.CellValue ParaIndentLeft { get; set; }
        public Core.CellValue ParaIndentRight { get; set; }
        public Core.CellValue ParaLocBulletFont { get; set; }
        public Core.CellValue ParaSpacingAfter { get; set; }
        public Core.CellValue ParaSpacingBefore { get; set; }
        public Core.CellValue ParaSpacingLine { get; set; }
        public Core.CellValue ParaTextPosAfterBullet { get; set; }

        //TextXForm
        public Core.CellValue TextXFormAngle { get; set; }
        public Core.CellValue TextXFormHeight { get; set; }
        public Core.CellValue TextXFormLocPinX { get; set; }
        public Core.CellValue TextXFormLocPinY { get; set; }
        public Core.CellValue TextXFormPinX { get; set; }
        public Core.CellValue TextXFormPinY { get; set; }
        public Core.CellValue TextXFormWidth { get; set; }

        // ShapeLayout
        public Core.CellValue ShapeLayoutConFixedCode { get; set; }
        public Core.CellValue ShapeLayoutConLineJumpCode { get; set; }
        public Core.CellValue ShapeLayoutConLineJumpDirX { get; set; }
        public Core.CellValue ShapeLayoutConLineJumpDirY { get; set; }
        public Core.CellValue ShapeLayoutConLineJumpStyle { get; set; }
        public Core.CellValue ShapeLayoutConLineRouteExt { get; set; }
        public Core.CellValue ShapeLayoutFixedCode { get; set; }
        public Core.CellValue ShapeLayoutPermeablePlace { get; set; }
        public Core.CellValue ShapeLayoutPermeableX { get; set; }
        public Core.CellValue ShapeLayoutPermeableY { get; set; }
        public Core.CellValue ShapeLayoutPlaceFlip { get; set; }
        public Core.CellValue ShapeLayoutPlaceStyle { get; set; }
        public Core.CellValue ShapeLayoutPlowCode { get; set; }
        public Core.CellValue ShapeLayoutRouteStyle { get; set; }
        public Core.CellValue ShapeLayoutSplit { get; set; }
        public Core.CellValue ShapeLayoutSplittable { get; set; }


        public void Apply(SidSrcWriter writer, short id)
        {
            writer.SetValue(id, Core.SrcConstants.XFormWidth, this.XFormWidth);
            writer.SetValue(id, Core.SrcConstants.XFormHeight, this.XFormHeight);
            writer.SetValue(id, Core.SrcConstants.XFormPinX, this.XFormPinX);
            writer.SetValue(id, Core.SrcConstants.XFormPinY, this.XFormPinY);
            writer.SetValue(id, Core.SrcConstants.XFormLocPinX, this.XFormLocPinX);
            writer.SetValue(id, Core.SrcConstants.XFormLocPinY, this.XFormLocPinY);
            writer.SetValue(id, Core.SrcConstants.XFormAngle, this.XFormAngle);
            writer.SetValue(id, Core.SrcConstants.LineBeginArrow, this.LineBeginArrow);
            writer.SetValue(id, Core.SrcConstants.LineBeginArrowSize, this.LineBeginArrowSize);

            writer.SetValue(id, Core.SrcConstants.FillBackground, this.FillBackground);
            writer.SetValue(id, Core.SrcConstants.FillBackgroundTransparency, this.FillBackgroundTransparency);
            writer.SetValue(id, Core.SrcConstants.FillForeground, this.FillForeground);
            writer.SetValue(id, Core.SrcConstants.FillForegroundTransparency, this.FillForegroundTransparency);
            writer.SetValue(id, Core.SrcConstants.FillPattern, this.FillPattern);

            writer.SetValue(id, Core.SrcConstants.FillShadowObliqueAngle, this.FillShadowObliqueAngle);
            writer.SetValue(id, Core.SrcConstants.FillShadowOffsetX, this.FillShadowOffsetX);
            writer.SetValue(id, Core.SrcConstants.FillShadowOffsetY, this.FillShadowOffsetY);
            writer.SetValue(id, Core.SrcConstants.FillShadowScaleFactor, this.FillShadowScaleFactor);
            writer.SetValue(id, Core.SrcConstants.FillShadowType, this.FillShadowType);
            writer.SetValue(id, Core.SrcConstants.FillShadowBackground, this.FillShadowBackground);
            writer.SetValue(id, Core.SrcConstants.FillShadowBackgroundTransparency, this.FillShadowBackgroundTransparency);
            writer.SetValue(id, Core.SrcConstants.FillShadowForeground, this.FillShadowForeground);
            writer.SetValue(id, Core.SrcConstants.FillShadowForegroundTransparency, this.FillShadowForegroundTransparency);
            writer.SetValue(id, Core.SrcConstants.FillShadowPattern, this.FillShadowPattern);

            writer.SetValue(id, Core.SrcConstants.CharCase, this.CharCase);
            writer.SetValue(id, Core.SrcConstants.CharFont, this.CharFont);
            writer.SetValue(id, Core.SrcConstants.CharColor, this.CharColor);
            writer.SetValue(id, Core.SrcConstants.CharSize, this.CharSize);
            writer.SetValue(id, Core.SrcConstants.CharLetterspace, this.CharLetterspace);
            writer.SetValue(id, Core.SrcConstants.CharStyle, this.CharStyle);
            writer.SetValue(id, Core.SrcConstants.CharColorTransparency, this.CharTransparency);

            writer.SetValue(id, Core.SrcConstants.LineEndArrow, this.LineEndArrow);
            writer.SetValue(id, Core.SrcConstants.LineEndArrowSize, this.LineEndArrowSize);

            // Line
            writer.SetValue(id, Core.SrcConstants.LineColor, this.LineColor);
            writer.SetValue(id, Core.SrcConstants.LineColorTransparency, this.LineColorTransparency);
            writer.SetValue(id, Core.SrcConstants.LinePattern, this.LinePattern);
            writer.SetValue(id, Core.SrcConstants.LineWeight, this.LineWeight);

            // Text
            writer.SetValue(id, Core.SrcConstants.TextBlockBottomMargin, this.TextBlockBottomMargin);
            writer.SetValue(id, Core.SrcConstants.TextBlockDefaultTabStop, this.TextBlockDefaultTabStop);
            writer.SetValue(id, Core.SrcConstants.TextBlockLeftMargin, this.TextBlockLeftMargin);
            writer.SetValue(id, Core.SrcConstants.TextBlockRightMargin, this.TextBlockRightMargin);
            writer.SetValue(id, Core.SrcConstants.TextBlockBackground, this.TextBlockBackground);
            writer.SetValue(id, Core.SrcConstants.TextBlockBackgroundTransparency, this.TextBlockBackgroundTransparency);
            writer.SetValue(id, Core.SrcConstants.TextBlockDirection, this.TextBlockDirection);
            writer.SetValue(id, Core.SrcConstants.TextBlockTopMargin, this.TextBlockTopMargin);
            writer.SetValue(id, Core.SrcConstants.TextBlockVerticalAlign, this.TextBlockVerticalAlign);

            // Paragraph

            writer.SetValue(id, Core.SrcConstants.ParaBullet, this.ParaBullet);
            writer.SetValue(id, Core.SrcConstants.ParaBulletFont, this.ParaBulletFont);
            writer.SetValue(id, Core.SrcConstants.ParaBulletFontSize, this.ParaBulletFontSize);
            writer.SetValue(id, Core.SrcConstants.ParaBulletString, this.ParaBulletString);
            writer.SetValue(id, Core.SrcConstants.ParaFlags, this.ParaFlags);
            writer.SetValue(id, Core.SrcConstants.ParaHorizontalAlign, this.ParaHorizontalAlign);
            writer.SetValue(id, Core.SrcConstants.ParaIndentFirst, this.ParaIndentFirst);
            writer.SetValue(id, Core.SrcConstants.ParaIndentLeft, this.ParaIndentLeft);
            writer.SetValue(id, Core.SrcConstants.ParaIndentRight, this.ParaIndentRight);
            writer.SetValue(id, Core.SrcConstants.ParaLocalizeBulletFont, this.ParaLocBulletFont);
            writer.SetValue(id, Core.SrcConstants.ParaSpacingAfter, this.ParaSpacingAfter);
            writer.SetValue(id, Core.SrcConstants.ParaSpacingBefore, this.ParaSpacingBefore);
            writer.SetValue(id, Core.SrcConstants.ParaSpacingLine, this.ParaSpacingLine);
            writer.SetValue(id, Core.SrcConstants.ParaTextPosAfterBullet, this.ParaTextPosAfterBullet);

            // TextXForm
            writer.SetValue(id, Core.SrcConstants.TextXFormAngle, this.TextXFormAngle);
            writer.SetValue(id, Core.SrcConstants.TextXFormHeight, this.TextXFormHeight);
            writer.SetValue(id, Core.SrcConstants.TextXFormLocPinX, this.TextXFormLocPinX);
            writer.SetValue(id, Core.SrcConstants.TextXFormLocPinY, this.TextXFormLocPinY);
            writer.SetValue(id, Core.SrcConstants.TextXFormPinX, this.TextXFormPinX);
            writer.SetValue(id, Core.SrcConstants.TextXFormPinY, this.TextXFormPinY);
            writer.SetValue(id, Core.SrcConstants.TextXFormWidth, this.TextXFormWidth);

            // ShapeLayout
            writer.SetValue(id, Core.SrcConstants.ShapeLayoutConnectorFixedCode, this.ShapeLayoutConFixedCode);
            writer.SetValue(id, Core.SrcConstants.ShapeLayoutLineJumpCode, this.ShapeLayoutConLineJumpCode);
            writer.SetValue(id, Core.SrcConstants.ShapeLayoutLineJumpDirX, this.ShapeLayoutConLineJumpDirX);
            writer.SetValue(id, Core.SrcConstants.ShapeLayoutLineJumpDirY, this.ShapeLayoutConLineJumpDirY);
            writer.SetValue(id, Core.SrcConstants.ShapeLayoutLineJumpStyle, this.ShapeLayoutConLineJumpStyle);
            writer.SetValue(id, Core.SrcConstants.ShapeLayoutLineRouteExt, this.ShapeLayoutConLineRouteExt);
            writer.SetValue(id, Core.SrcConstants.ShapeLayoutShapeFixedCode, this.ShapeLayoutFixedCode);
            writer.SetValue(id, Core.SrcConstants.ShapeLayoutShapePermeablePlace, this.ShapeLayoutPermeablePlace);
            writer.SetValue(id, Core.SrcConstants.ShapeLayoutShapePermeableX, this.ShapeLayoutPermeableX);
            writer.SetValue(id, Core.SrcConstants.ShapeLayoutShapePermeableY, this.ShapeLayoutPermeableY);
            writer.SetValue(id, Core.SrcConstants.ShapeLayoutShapePlaceFlip, this.ShapeLayoutPlaceFlip);
            writer.SetValue(id, Core.SrcConstants.ShapeLayoutShapePlaceStyle, this.ShapeLayoutPlaceStyle);
            writer.SetValue(id, Core.SrcConstants.ShapeLayoutShapePlowCode, this.ShapeLayoutPlowCode);
            writer.SetValue(id, Core.SrcConstants.ShapeLayoutShapeRouteStyle, this.ShapeLayoutRouteStyle);
            writer.SetValue(id, Core.SrcConstants.ShapeLayoutShapeSplit, this.ShapeLayoutSplit);
            writer.SetValue(id, Core.SrcConstants.ShapeLayoutShapeSplittable, this.ShapeLayoutSplittable);
        }

        public ShapeCells ShallowCopy()
        {
            return (ShapeCells) this.MemberwiseClone();
        }

        public void ApplyFormulasTo(ShapeCells target)
        {
            if (this.XFormWidth.HasValue) { target.XFormWidth = this.XFormWidth; }
            if (this.XFormHeight.HasValue) { target.XFormHeight = this.XFormHeight; }
            if (this.XFormAngle.HasValue) { target.XFormAngle = this.XFormAngle; }
            if (this.XFormPinX.HasValue) { target.XFormPinX = this.XFormPinX; }
            if (this.XFormPinY.HasValue) { target.XFormPinY = this.XFormPinY; }
            if (this.XFormLocPinX.HasValue) { target.XFormLocPinX = this.XFormLocPinX; }
            if (this.XFormLocPinY.HasValue) { target.XFormLocPinY = this.XFormLocPinY; }
            if (this.FillBackground.HasValue) { target.FillBackground = this.FillBackground; }
            if (this.FillBackgroundTransparency.HasValue) { target.FillBackgroundTransparency = this.FillBackgroundTransparency; }
            if (this.FillForeground.HasValue) { target.FillForeground = this.FillForeground; }
            if (this.FillForegroundTransparency.HasValue) { target.FillForegroundTransparency = this.FillForegroundTransparency; }
            if (this.FillPattern.HasValue) { target.FillPattern = this.FillPattern; }
            if (this.FillShadowObliqueAngle.HasValue) { target.FillShadowObliqueAngle = this.FillShadowObliqueAngle; }
            if (this.FillShadowOffsetX.HasValue) { target.FillShadowOffsetX = this.FillShadowOffsetX; }
            if (this.FillShadowOffsetY.HasValue) { target.FillShadowOffsetY = this.FillShadowOffsetY; }
            if (this.FillShadowScaleFactor.HasValue) { target.FillShadowScaleFactor = this.FillShadowScaleFactor; }
            if (this.FillShadowType.HasValue) { target.FillShadowType = this.FillShadowType; }
            if (this.FillShadowBackground.HasValue) { target.FillShadowBackground = this.FillShadowBackground; }
            if (this.FillShadowBackgroundTransparency.HasValue) { target.FillShadowBackgroundTransparency = this.FillShadowBackgroundTransparency; }
            if (this.FillShadowForeground.HasValue) { target.FillShadowForeground = this.FillShadowForeground; }
            if (this.FillShadowForegroundTransparency.HasValue) { target.FillShadowForegroundTransparency = this.FillShadowForegroundTransparency; }
            if (this.FillShadowPattern.HasValue) { target.FillShadowPattern = this.FillShadowPattern; }
            if (this.LineBeginArrow.HasValue) { target.LineBeginArrow = this.LineBeginArrow; }
            if (this.LineBeginArrowSize.HasValue) { target.LineBeginArrowSize = this.LineBeginArrowSize; }
            if (this.LineEndArrow.HasValue) { target.LineEndArrow = this.LineEndArrow; }
            if (this.LineEndArrowSize.HasValue) { target.LineEndArrowSize = this.LineEndArrowSize; }
            if (this.LineCap.HasValue) { target.LineCap = this.LineCap; }
            if (this.LineColor.HasValue) { target.LineColor = this.LineColor; }
            if (this.LineColorTransparency.HasValue) { target.LineColorTransparency = this.LineColorTransparency; }
            if (this.LinePattern.HasValue) { target.LinePattern = this.LinePattern; }
            if (this.LineWeight.HasValue) { target.LineWeight = this.LineWeight; }
            if (this.Rounding.HasValue) { target.Rounding = this.Rounding; }
            if (this.CharAsianFont.HasValue) { target.CharAsianFont = this.CharAsianFont; }
            if (this.CharCase.HasValue) { target.CharCase = this.CharCase; }
            if (this.CharColor.HasValue) { target.CharColor = this.CharColor; }
            if (this.CharComplexScriptFont.HasValue) { target.CharComplexScriptFont = this.CharComplexScriptFont; }
            if (this.CharComplexScriptSize.HasValue) { target.CharComplexScriptSize = this.CharComplexScriptSize; }
            if (this.CharDoubleStrikeThrough.HasValue) { target.CharDoubleStrikeThrough = this.CharDoubleStrikeThrough; }
            if (this.CharDoubleUnderline.HasValue) { target.CharDoubleUnderline = this.CharDoubleUnderline; }
            if (this.CharFont.HasValue) { target.CharFont = this.CharFont; }
            if (this.CharLangID.HasValue) { target.CharLangID = this.CharLangID; }
            if (this.CharLocale.HasValue) { target.CharLocale = this.CharLocale; }
            if (this.CharLocalizeFont.HasValue) { target.CharLocalizeFont = this.CharLocalizeFont; }
            if (this.CharOverline.HasValue) { target.CharOverline = this.CharOverline; }
            if (this.CharPerpendicular.HasValue) { target.CharPerpendicular = this.CharPerpendicular; }
            if (this.CharPos.HasValue) { target.CharPos = this.CharPos; }
            if (this.CharRTLText.HasValue) { target.CharRTLText = this.CharRTLText; }
            if (this.CharFontScale.HasValue) { target.CharFontScale = this.CharFontScale; }
            if (this.CharSize.HasValue) { target.CharSize = this.CharSize; }
            if (this.CharLetterspace.HasValue) { target.CharLetterspace = this.CharLetterspace; }
            if (this.CharStrikethru.HasValue) { target.CharStrikethru = this.CharStrikethru; }
            if (this.CharStyle.HasValue) { target.CharStyle = this.CharStyle; }
            if (this.CharTransparency.HasValue) { target.CharTransparency = this.CharTransparency; }
            if (this.CharUseVertical.HasValue) { target.CharUseVertical = this.CharUseVertical; }
            if (this.TextBlockBottomMargin.HasValue) { target.TextBlockBottomMargin = this.TextBlockBottomMargin; }
            if (this.TextBlockDefaultTabStop.HasValue) { target.TextBlockDefaultTabStop = this.TextBlockDefaultTabStop; }
            if (this.TextBlockLeftMargin.HasValue) { target.TextBlockLeftMargin = this.TextBlockLeftMargin; }
            if (this.TextBlockRightMargin.HasValue) { target.TextBlockRightMargin = this.TextBlockRightMargin; }
            if (this.TextBlockBackground.HasValue) { target.TextBlockBackground = this.TextBlockBackground; }
            if (this.TextBlockBackgroundTransparency.HasValue) { target.TextBlockBackgroundTransparency = this.TextBlockBackgroundTransparency; }
            if (this.TextBlockDirection.HasValue) { target.TextBlockDirection = this.TextBlockDirection; }
            if (this.TextBlockTopMargin.HasValue) { target.TextBlockTopMargin = this.TextBlockTopMargin; }
            if (this.TextBlockVerticalAlign.HasValue) { target.TextBlockVerticalAlign = this.TextBlockVerticalAlign; }
            if (this.ParaBullet.HasValue) { target.ParaBullet = this.ParaBullet; }
            if (this.ParaBulletFont.HasValue) { target.ParaBulletFont = this.ParaBulletFont; }
            if (this.ParaBulletFontSize.HasValue) { target.ParaBulletFontSize = this.ParaBulletFontSize; }
            if (this.ParaBulletString.HasValue) { target.ParaBulletString = this.ParaBulletString; }
            if (this.ParaFlags.HasValue) { target.ParaFlags = this.ParaFlags; }
            if (this.ParaHorizontalAlign.HasValue) { target.ParaHorizontalAlign = this.ParaHorizontalAlign; }
            if (this.ParaIndentFirst.HasValue) { target.ParaIndentFirst = this.ParaIndentFirst; }
            if (this.ParaIndentLeft.HasValue) { target.ParaIndentLeft = this.ParaIndentLeft; }
            if (this.ParaIndentRight.HasValue) { target.ParaIndentRight = this.ParaIndentRight; }
            if (this.ParaLocBulletFont.HasValue) { target.ParaLocBulletFont = this.ParaLocBulletFont; }
            if (this.ParaSpacingAfter.HasValue) { target.ParaSpacingAfter = this.ParaSpacingAfter; }
            if (this.ParaSpacingBefore.HasValue) { target.ParaSpacingBefore = this.ParaSpacingBefore; }
            if (this.ParaSpacingLine.HasValue) { target.ParaSpacingLine = this.ParaSpacingLine; }
            if (this.ParaTextPosAfterBullet.HasValue) { target.ParaTextPosAfterBullet = this.ParaTextPosAfterBullet; }
            if (this.TextXFormAngle.HasValue) { target.TextXFormAngle = this.TextXFormAngle; }
            if (this.TextXFormHeight.HasValue) { target.TextXFormHeight = this.TextXFormHeight; }
            if (this.TextXFormLocPinX.HasValue) { target.TextXFormLocPinX = this.TextXFormLocPinX; }
            if (this.TextXFormLocPinY.HasValue) { target.TextXFormLocPinY = this.TextXFormLocPinY; }
            if (this.TextXFormPinX.HasValue) { target.TextXFormPinX = this.TextXFormPinX; }
            if (this.TextXFormPinY.HasValue) { target.TextXFormPinY = this.TextXFormPinY; }
            if (this.TextXFormWidth.HasValue) { target.TextXFormWidth = this.TextXFormWidth; }
            if (this.ShapeLayoutConFixedCode.HasValue) { target.ShapeLayoutConFixedCode = this.ShapeLayoutConFixedCode; }
            if (this.ShapeLayoutConLineJumpCode.HasValue) { target.ShapeLayoutConLineJumpCode = this.ShapeLayoutConLineJumpCode; }
            if (this.ShapeLayoutConLineJumpDirX.HasValue) { target.ShapeLayoutConLineJumpDirX = this.ShapeLayoutConLineJumpDirX; }
            if (this.ShapeLayoutConLineJumpDirY.HasValue) { target.ShapeLayoutConLineJumpDirY = this.ShapeLayoutConLineJumpDirY; }
            if (this.ShapeLayoutConLineJumpStyle.HasValue) { target.ShapeLayoutConLineJumpStyle = this.ShapeLayoutConLineJumpStyle; }
            if (this.ShapeLayoutConLineRouteExt.HasValue) { target.ShapeLayoutConLineRouteExt = this.ShapeLayoutConLineRouteExt; }
            if (this.ShapeLayoutFixedCode.HasValue) { target.ShapeLayoutFixedCode = this.ShapeLayoutFixedCode; }
            if (this.ShapeLayoutPermeablePlace.HasValue) { target.ShapeLayoutPermeablePlace = this.ShapeLayoutPermeablePlace; }
            if (this.ShapeLayoutPermeableX.HasValue) { target.ShapeLayoutPermeableX = this.ShapeLayoutPermeableX; }
            if (this.ShapeLayoutPermeableY.HasValue) { target.ShapeLayoutPermeableY = this.ShapeLayoutPermeableY; }
            if (this.ShapeLayoutPlaceFlip.HasValue) { target.ShapeLayoutPlaceFlip = this.ShapeLayoutPlaceFlip; }
            if (this.ShapeLayoutPlaceStyle.HasValue) { target.ShapeLayoutPlaceStyle = this.ShapeLayoutPlaceStyle; }
            if (this.ShapeLayoutPlowCode.HasValue) { target.ShapeLayoutPlowCode = this.ShapeLayoutPlowCode; }
            if (this.ShapeLayoutRouteStyle.HasValue) { target.ShapeLayoutRouteStyle = this.ShapeLayoutRouteStyle; }
            if (this.ShapeLayoutSplit.HasValue) { target.ShapeLayoutSplit = this.ShapeLayoutSplit; }
            if (this.ShapeLayoutSplittable.HasValue) { target.ShapeLayoutSplittable = this.ShapeLayoutSplittable; }
        }
    }
}