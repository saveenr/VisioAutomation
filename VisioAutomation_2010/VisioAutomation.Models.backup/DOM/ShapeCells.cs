using VisioAutomation.ShapeSheet.Writers;

namespace VisioAutomation.Models.Dom
{
    public class ShapeCells
    {
        // XFormOut
        public ShapeSheet.CellValue XFormWidth { get; set; }
        public ShapeSheet.CellValue XFormHeight { get; set; }
        public ShapeSheet.CellValue XFormAngle { get; set; }
        public ShapeSheet.CellValue XFormPinX { get; set; }
        public ShapeSheet.CellValue XFormPinY { get; set; }
        public ShapeSheet.CellValue XFormLocPinX { get; set; }
        public ShapeSheet.CellValue XFormLocPinY { get; set; }

        // Fill
        public ShapeSheet.CellValue FillBackground { get; set; }
        public ShapeSheet.CellValue FillBackgroundTransparency { get; set; }
        public ShapeSheet.CellValue FillForeground { get; set; }
        public ShapeSheet.CellValue FillForegroundTransparency { get; set; }
        public ShapeSheet.CellValue FillPattern { get; set; }
        public ShapeSheet.CellValue FillShadowObliqueAngle { get; set; }
        public ShapeSheet.CellValue FillShadowOffsetX { get; set; }
        public ShapeSheet.CellValue FillShadowOffsetY { get; set; }
        public ShapeSheet.CellValue FillShadowScaleFactor { get; set; }
        public ShapeSheet.CellValue FillShadowType { get; set; }
        public ShapeSheet.CellValue FillShadowBackground { get; set; }
        public ShapeSheet.CellValue FillShadowBackgroundTransparency { get; set; }
        public ShapeSheet.CellValue FillShadowForeground { get; set; }
        public ShapeSheet.CellValue FillShadowForegroundTransparency { get; set; }
        public ShapeSheet.CellValue FillShadowPattern { get; set; }

        // Line
        public ShapeSheet.CellValue LineBeginArrow { get; set; }
        public ShapeSheet.CellValue LineBeginArrowSize { get; set; }
        public ShapeSheet.CellValue LineEndArrow { get; set; }
        public ShapeSheet.CellValue LineEndArrowSize { get; set; }
        public ShapeSheet.CellValue LineCap { get; set; }
        public ShapeSheet.CellValue LineColor { get; set; }
        public ShapeSheet.CellValue LineColorTransparency { get; set; }
        public ShapeSheet.CellValue LinePattern { get; set; }
        public ShapeSheet.CellValue LineWeight { get; set; }
        public ShapeSheet.CellValue Rounding { get; set; }

        // Character
        public ShapeSheet.CellValue CharAsianFont { get; set; }
        public ShapeSheet.CellValue CharCase { get; set; }
        public ShapeSheet.CellValue CharColor { get; set; }
        public ShapeSheet.CellValue CharComplexScriptFont { get; set; }
        public ShapeSheet.CellValue CharComplexScriptSize { get; set; }
        public ShapeSheet.CellValue CharDoubleStrikeThrough { get; set; }
        public ShapeSheet.CellValue CharDoubleUnderline { get; set; }
        public ShapeSheet.CellValue CharFont { get; set; }
        public ShapeSheet.CellValue CharFontScale { get; set; }
        public ShapeSheet.CellValue CharLangID { get; set; }
        public ShapeSheet.CellValue CharLetterspace { get; set; }
        public ShapeSheet.CellValue CharLocale { get; set; }
        public ShapeSheet.CellValue CharLocalizeFont { get; set; }
        public ShapeSheet.CellValue CharOverline { get; set; }
        public ShapeSheet.CellValue CharPerpendicular { get; set; }
        public ShapeSheet.CellValue CharPos { get; set; }
        public ShapeSheet.CellValue CharRTLText { get; set; }
        public ShapeSheet.CellValue CharSize { get; set; }
        public ShapeSheet.CellValue CharStrikethru { get; set; }
        public ShapeSheet.CellValue CharStyle { get; set; }
        public ShapeSheet.CellValue CharTransparency { get; set; }
        public ShapeSheet.CellValue CharUseVertical { get; set; }

        // Text Block
        public ShapeSheet.CellValue TextBlockBottomMargin { get; set; }
        public ShapeSheet.CellValue TextBlockDefaultTabStop { get; set; }
        public ShapeSheet.CellValue TextBlockLeftMargin { get; set; }
        public ShapeSheet.CellValue TextBlockRightMargin { get; set; }
        public ShapeSheet.CellValue TextBlockBackground { get; set; }
        public ShapeSheet.CellValue TextBlockBackgroundTransparency { get; set; }
        public ShapeSheet.CellValue TextBlockDirection { get; set; }
        public ShapeSheet.CellValue TextBlockTopMargin { get; set; }
        public ShapeSheet.CellValue TextBlockVerticalAlign { get; set; }

        // Paragraph
        public ShapeSheet.CellValue ParaBullet { get; set; }
        public ShapeSheet.CellValue ParaBulletFont { get; set; }
        public ShapeSheet.CellValue ParaBulletFontSize { get; set; }
        public ShapeSheet.CellValue ParaBulletString { get; set; }
        public ShapeSheet.CellValue ParaFlags { get; set; }
        public ShapeSheet.CellValue ParaHorizontalAlign { get; set; }
        public ShapeSheet.CellValue ParaIndentFirst { get; set; }
        public ShapeSheet.CellValue ParaIndentLeft { get; set; }
        public ShapeSheet.CellValue ParaIndentRight { get; set; }
        public ShapeSheet.CellValue ParaLocBulletFont { get; set; }
        public ShapeSheet.CellValue ParaSpacingAfter { get; set; }
        public ShapeSheet.CellValue ParaSpacingBefore { get; set; }
        public ShapeSheet.CellValue ParaSpacingLine { get; set; }
        public ShapeSheet.CellValue ParaTextPosAfterBullet { get; set; }

        //TextXForm
        public ShapeSheet.CellValue TextXFormAngle { get; set; }
        public ShapeSheet.CellValue TextXFormHeight { get; set; }
        public ShapeSheet.CellValue TextXFormLocPinX { get; set; }
        public ShapeSheet.CellValue TextXFormLocPinY { get; set; }
        public ShapeSheet.CellValue TextXFormPinX { get; set; }
        public ShapeSheet.CellValue TextXFormPinY { get; set; }
        public ShapeSheet.CellValue TextXFormWidth { get; set; }

        // ShapeLayout
        public ShapeSheet.CellValue ShapeLayoutConFixedCode { get; set; }
        public ShapeSheet.CellValue ShapeLayoutConLineJumpCode { get; set; }
        public ShapeSheet.CellValue ShapeLayoutConLineJumpDirX { get; set; }
        public ShapeSheet.CellValue ShapeLayoutConLineJumpDirY { get; set; }
        public ShapeSheet.CellValue ShapeLayoutConLineJumpStyle { get; set; }
        public ShapeSheet.CellValue ShapeLayoutConLineRouteExt { get; set; }
        public ShapeSheet.CellValue ShapeLayoutFixedCode { get; set; }
        public ShapeSheet.CellValue ShapeLayoutPermeablePlace { get; set; }
        public ShapeSheet.CellValue ShapeLayoutPermeableX { get; set; }
        public ShapeSheet.CellValue ShapeLayoutPermeableY { get; set; }
        public ShapeSheet.CellValue ShapeLayoutPlaceFlip { get; set; }
        public ShapeSheet.CellValue ShapeLayoutPlaceStyle { get; set; }
        public ShapeSheet.CellValue ShapeLayoutPlowCode { get; set; }
        public ShapeSheet.CellValue ShapeLayoutRouteStyle { get; set; }
        public ShapeSheet.CellValue ShapeLayoutSplit { get; set; }
        public ShapeSheet.CellValue ShapeLayoutSplittable { get; set; }


        public void Apply(SidSrcWriter writer, short id)
        {
            writer.SetValue(id, ShapeSheet.SrcConstants.XFormWidth, this.XFormWidth);
            writer.SetValue(id, ShapeSheet.SrcConstants.XFormHeight, this.XFormHeight);
            writer.SetValue(id, ShapeSheet.SrcConstants.XFormPinX, this.XFormPinX);
            writer.SetValue(id, ShapeSheet.SrcConstants.XFormPinY, this.XFormPinY);
            writer.SetValue(id, ShapeSheet.SrcConstants.XFormLocPinX, this.XFormLocPinX);
            writer.SetValue(id, ShapeSheet.SrcConstants.XFormLocPinY, this.XFormLocPinY);
            writer.SetValue(id, ShapeSheet.SrcConstants.XFormAngle, this.XFormAngle);
            writer.SetValue(id, ShapeSheet.SrcConstants.LineBeginArrow, this.LineBeginArrow);
            writer.SetValue(id, ShapeSheet.SrcConstants.LineBeginArrowSize, this.LineBeginArrowSize);

            writer.SetValue(id, ShapeSheet.SrcConstants.FillBackground, this.FillBackground);
            writer.SetValue(id, ShapeSheet.SrcConstants.FillBackgroundTransparency, this.FillBackgroundTransparency);
            writer.SetValue(id, ShapeSheet.SrcConstants.FillForeground, this.FillForeground);
            writer.SetValue(id, ShapeSheet.SrcConstants.FillForegroundTransparency, this.FillForegroundTransparency);
            writer.SetValue(id, ShapeSheet.SrcConstants.FillPattern, this.FillPattern);

            writer.SetValue(id, ShapeSheet.SrcConstants.FillShadowObliqueAngle, this.FillShadowObliqueAngle);
            writer.SetValue(id, ShapeSheet.SrcConstants.FillShadowOffsetX, this.FillShadowOffsetX);
            writer.SetValue(id, ShapeSheet.SrcConstants.FillShadowOffsetY, this.FillShadowOffsetY);
            writer.SetValue(id, ShapeSheet.SrcConstants.FillShadowScaleFactor, this.FillShadowScaleFactor);
            writer.SetValue(id, ShapeSheet.SrcConstants.FillShadowType, this.FillShadowType);
            writer.SetValue(id, ShapeSheet.SrcConstants.FillShadowBackground, this.FillShadowBackground);
            writer.SetValue(id, ShapeSheet.SrcConstants.FillShadowBackgroundTransparency, this.FillShadowBackgroundTransparency);
            writer.SetValue(id, ShapeSheet.SrcConstants.FillShadowForeground, this.FillShadowForeground);
            writer.SetValue(id, ShapeSheet.SrcConstants.FillShadowForegroundTransparency, this.FillShadowForegroundTransparency);
            writer.SetValue(id, ShapeSheet.SrcConstants.FillShadowPattern, this.FillShadowPattern);

            writer.SetValue(id, ShapeSheet.SrcConstants.CharCase, this.CharCase);
            writer.SetValue(id, ShapeSheet.SrcConstants.CharFont, this.CharFont);
            writer.SetValue(id, ShapeSheet.SrcConstants.CharColor, this.CharColor);
            writer.SetValue(id, ShapeSheet.SrcConstants.CharSize, this.CharSize);
            writer.SetValue(id, ShapeSheet.SrcConstants.CharLetterspace, this.CharLetterspace);
            writer.SetValue(id, ShapeSheet.SrcConstants.CharStyle, this.CharStyle);
            writer.SetValue(id, ShapeSheet.SrcConstants.CharColorTransparency, this.CharTransparency);

            writer.SetValue(id, ShapeSheet.SrcConstants.LineEndArrow, this.LineEndArrow);
            writer.SetValue(id, ShapeSheet.SrcConstants.LineEndArrowSize, this.LineEndArrowSize);

            // Line
            writer.SetValue(id, ShapeSheet.SrcConstants.LineColor, this.LineColor);
            writer.SetValue(id, ShapeSheet.SrcConstants.LineColorTransparency, this.LineColorTransparency);
            writer.SetValue(id, ShapeSheet.SrcConstants.LinePattern, this.LinePattern);
            writer.SetValue(id, ShapeSheet.SrcConstants.LineWeight, this.LineWeight);

            // Text
            writer.SetValue(id, ShapeSheet.SrcConstants.TextBlockBottomMargin, this.TextBlockBottomMargin);
            writer.SetValue(id, ShapeSheet.SrcConstants.TextBlockDefaultTabStop, this.TextBlockDefaultTabStop);
            writer.SetValue(id, ShapeSheet.SrcConstants.TextBlockLeftMargin, this.TextBlockLeftMargin);
            writer.SetValue(id, ShapeSheet.SrcConstants.TextBlockRightMargin, this.TextBlockRightMargin);
            writer.SetValue(id, ShapeSheet.SrcConstants.TextBlockBackground, this.TextBlockBackground);
            writer.SetValue(id, ShapeSheet.SrcConstants.TextBlockBackgroundTransparency, this.TextBlockBackgroundTransparency);
            writer.SetValue(id, ShapeSheet.SrcConstants.TextBlockDirection, this.TextBlockDirection);
            writer.SetValue(id, ShapeSheet.SrcConstants.TextBlockTopMargin, this.TextBlockTopMargin);
            writer.SetValue(id, ShapeSheet.SrcConstants.TextBlockVerticalAlign, this.TextBlockVerticalAlign);

            // Paragraph

            writer.SetValue(id, ShapeSheet.SrcConstants.ParaBullet, this.ParaBullet);
            writer.SetValue(id, ShapeSheet.SrcConstants.ParaBulletFont, this.ParaBulletFont);
            writer.SetValue(id, ShapeSheet.SrcConstants.ParaBulletFontSize, this.ParaBulletFontSize);
            writer.SetValue(id, ShapeSheet.SrcConstants.ParaBulletString, this.ParaBulletString);
            writer.SetValue(id, ShapeSheet.SrcConstants.ParaFlags, this.ParaFlags);
            writer.SetValue(id, ShapeSheet.SrcConstants.ParaHorizontalAlign, this.ParaHorizontalAlign);
            writer.SetValue(id, ShapeSheet.SrcConstants.ParaIndentFirst, this.ParaIndentFirst);
            writer.SetValue(id, ShapeSheet.SrcConstants.ParaIndentLeft, this.ParaIndentLeft);
            writer.SetValue(id, ShapeSheet.SrcConstants.ParaIndentRight, this.ParaIndentRight);
            writer.SetValue(id, ShapeSheet.SrcConstants.ParaLocalizeBulletFont, this.ParaLocBulletFont);
            writer.SetValue(id, ShapeSheet.SrcConstants.ParaSpacingAfter, this.ParaSpacingAfter);
            writer.SetValue(id, ShapeSheet.SrcConstants.ParaSpacingBefore, this.ParaSpacingBefore);
            writer.SetValue(id, ShapeSheet.SrcConstants.ParaSpacingLine, this.ParaSpacingLine);
            writer.SetValue(id, ShapeSheet.SrcConstants.ParaTextPosAfterBullet, this.ParaTextPosAfterBullet);

            // TextXForm
            writer.SetValue(id, ShapeSheet.SrcConstants.TextXFormAngle, this.TextXFormAngle);
            writer.SetValue(id, ShapeSheet.SrcConstants.TextXFormHeight, this.TextXFormHeight);
            writer.SetValue(id, ShapeSheet.SrcConstants.TextXFormLocPinX, this.TextXFormLocPinX);
            writer.SetValue(id, ShapeSheet.SrcConstants.TextXFormLocPinY, this.TextXFormLocPinY);
            writer.SetValue(id, ShapeSheet.SrcConstants.TextXFormPinX, this.TextXFormPinX);
            writer.SetValue(id, ShapeSheet.SrcConstants.TextXFormPinY, this.TextXFormPinY);
            writer.SetValue(id, ShapeSheet.SrcConstants.TextXFormWidth, this.TextXFormWidth);

            // ShapeLayout
            writer.SetValue(id, ShapeSheet.SrcConstants.ShapeLayoutConnectorFixedCode, this.ShapeLayoutConFixedCode);
            writer.SetValue(id, ShapeSheet.SrcConstants.ShapeLayoutLineJumpCode, this.ShapeLayoutConLineJumpCode);
            writer.SetValue(id, ShapeSheet.SrcConstants.ShapeLayoutLineJumpDirX, this.ShapeLayoutConLineJumpDirX);
            writer.SetValue(id, ShapeSheet.SrcConstants.ShapeLayoutLineJumpDirY, this.ShapeLayoutConLineJumpDirY);
            writer.SetValue(id, ShapeSheet.SrcConstants.ShapeLayoutLineJumpStyle, this.ShapeLayoutConLineJumpStyle);
            writer.SetValue(id, ShapeSheet.SrcConstants.ShapeLayoutLineRouteExt, this.ShapeLayoutConLineRouteExt);
            writer.SetValue(id, ShapeSheet.SrcConstants.ShapeLayoutShapeFixedCode, this.ShapeLayoutFixedCode);
            writer.SetValue(id, ShapeSheet.SrcConstants.ShapeLayoutShapePermeablePlace, this.ShapeLayoutPermeablePlace);
            writer.SetValue(id, ShapeSheet.SrcConstants.ShapeLayoutShapePermeableX, this.ShapeLayoutPermeableX);
            writer.SetValue(id, ShapeSheet.SrcConstants.ShapeLayoutShapePermeableY, this.ShapeLayoutPermeableY);
            writer.SetValue(id, ShapeSheet.SrcConstants.ShapeLayoutShapePlaceFlip, this.ShapeLayoutPlaceFlip);
            writer.SetValue(id, ShapeSheet.SrcConstants.ShapeLayoutShapePlaceStyle, this.ShapeLayoutPlaceStyle);
            writer.SetValue(id, ShapeSheet.SrcConstants.ShapeLayoutShapePlowCode, this.ShapeLayoutPlowCode);
            writer.SetValue(id, ShapeSheet.SrcConstants.ShapeLayoutShapeRouteStyle, this.ShapeLayoutRouteStyle);
            writer.SetValue(id, ShapeSheet.SrcConstants.ShapeLayoutShapeSplit, this.ShapeLayoutSplit);
            writer.SetValue(id, ShapeSheet.SrcConstants.ShapeLayoutShapeSplittable, this.ShapeLayoutSplittable);
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