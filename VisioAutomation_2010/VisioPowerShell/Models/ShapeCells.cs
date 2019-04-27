using System.Collections.Generic;
using SRCCON = VisioAutomation.ShapeSheet.SrcConstants;

namespace VisioPowerShell.Models
{
    public class ShapeCells : VisioPowerShell.Models.BaseCells
    {
        // Shape XForm
        public string XFormAngle;
        public string XFormHeight;
        public string XFormLocPinX;
        public string XFormLocPinY;
        public string XFormPinX;
        public string XFormPinY;
        public string XFormWidth;

        // Shape Fill
        public string FillBackground;
        public string FillBackgroundTransparency;
        public string FillForeground;
        public string FillForegroundTransparency;
        public string FillPattern;
        public string FillShadowBackground;
        public string FillShadowBackgroundTransparency;
        public string FillShadowForeground;
        public string FillShadowForegroundTransparency;
        public string FillShadowPattern;

        // Shape Line

        public string LineBeginArrow;
        public string LineBeginArrowSize;
        public string LineCap;
        public string LineColor;
        public string LineEndArrow;
        public string LineEndArrowSize;
        public string LinePattern;
        public string LineRounding;
        public string LineWeight;

        // Shape 1-D
        public string OneDBeginX;
        public string OneDBeginY;
        public string OneDEndX;
        public string OneDEndY;

        // Shape Other
        public string GroupSelectMode;

        // Shape Character
        public string CharCase;
        public string CharColor;
        public string CharColorTransparency;
        public string CharFont;
        public string CharFontScale;
        public string CharLetterspace;
        public string CharSize;
        public string CharStyle;
        public string CharDoubleStrikethrough { get; set; }
        public string CharDoubleUnderline { get; set; }
        public string CharLangID { get; set; }
        public string CharLocale { get; set; }
        public string CharOverline { get; set; }
        public string CharPerpendicular { get; set; }
        public string CharPos { get; set; }
        public string CharStrikethru { get; set; }

        // Shape TextXForm
        public string TextFormAngle;
        public string TextFormHeight;
        public string TextFormLocPinX;
        public string TextFormLocPinY;
        public string TextFormPinX;
        public string TextFormPinY;
        public string TextFormWidth;

        internal override IEnumerable<Internal.CellTuple> EnumCellTuples()
        {
            // Shape XForm
            yield return new Internal.CellTuple(nameof(SRCCON.XFormAngle), SRCCON.XFormAngle, this.XFormAngle);
            yield return new Internal.CellTuple(nameof(SRCCON.XFormHeight), SRCCON.XFormHeight, this.XFormHeight);
            yield return new Internal.CellTuple(nameof(SRCCON.XFormLocPinX), SRCCON.XFormLocPinX, this.XFormLocPinX);
            yield return new Internal.CellTuple(nameof(SRCCON.XFormLocPinY), SRCCON.XFormLocPinY, this.XFormLocPinY);
            yield return new Internal.CellTuple(nameof(SRCCON.XFormPinX), SRCCON.XFormPinX, this.XFormPinX);
            yield return new Internal.CellTuple(nameof(SRCCON.XFormPinY), SRCCON.XFormPinY, this.XFormPinY);
            yield return new Internal.CellTuple(nameof(SRCCON.XFormWidth), SRCCON.XFormWidth, this.XFormWidth);

            // Shape Fill
            yield return new Internal.CellTuple(nameof(SRCCON.FillBackground), SRCCON.FillBackground, this.FillBackground);
            yield return new Internal.CellTuple(nameof(SRCCON.FillBackgroundTransparency), SRCCON.FillBackgroundTransparency,
                this.FillBackgroundTransparency);
            yield return new Internal.CellTuple(nameof(SRCCON.FillForeground), SRCCON.FillForeground, this.FillForeground);
            yield return new Internal.CellTuple(nameof(SRCCON.FillForegroundTransparency), SRCCON.FillForegroundTransparency,
                this.FillForegroundTransparency);
            yield return new Internal.CellTuple(nameof(SRCCON.FillPattern), SRCCON.FillPattern, this.FillPattern);
            yield return new Internal.CellTuple(nameof(SRCCON.FillShadowBackground), SRCCON.FillShadowBackground,
                this.FillShadowBackground);
            yield return new Internal.CellTuple(nameof(SRCCON.FillShadowBackgroundTransparency),
                SRCCON.FillShadowBackgroundTransparency, this.FillShadowBackgroundTransparency);
            yield return new Internal.CellTuple(nameof(SRCCON.FillShadowForeground), SRCCON.FillShadowForeground,
                this.FillShadowForeground);
            yield return new Internal.CellTuple(nameof(SRCCON.FillShadowForegroundTransparency),
                SRCCON.FillShadowForegroundTransparency, this.FillShadowForegroundTransparency);
            yield return new Internal.CellTuple(nameof(SRCCON.FillShadowPattern), SRCCON.FillShadowPattern,
                this.FillShadowPattern);

            // Shape Line
            yield return new Internal.CellTuple(nameof(SRCCON.LineBeginArrow), SRCCON.LineBeginArrow, this.LineBeginArrow);
            yield return new Internal.CellTuple(nameof(SRCCON.LineBeginArrowSize), SRCCON.LineBeginArrowSize,
                this.LineBeginArrowSize);
            yield return new Internal.CellTuple(nameof(SRCCON.LineCap), SRCCON.LineCap, this.LineCap);
            yield return new Internal.CellTuple(nameof(SRCCON.LineColor), SRCCON.LineColor, this.LineColor);
            yield return new Internal.CellTuple(nameof(SRCCON.LineEndArrow), SRCCON.LineEndArrow, this.LineEndArrow);
            yield return new Internal.CellTuple(nameof(SRCCON.LineEndArrowSize), SRCCON.LineEndArrowSize, this.LineEndArrowSize);
            yield return new Internal.CellTuple(nameof(SRCCON.LinePattern), SRCCON.LinePattern, this.LinePattern);
            yield return new Internal.CellTuple(nameof(SRCCON.LineRounding), SRCCON.LineRounding, this.LineRounding);
            yield return new Internal.CellTuple(nameof(SRCCON.LineWeight), SRCCON.LineWeight, this.LineWeight);

            // Shape 1-D
            yield return new Internal.CellTuple(nameof(SRCCON.OneDBeginX), SRCCON.OneDBeginX, this.OneDBeginX);
            yield return new Internal.CellTuple(nameof(SRCCON.OneDBeginY), SRCCON.OneDBeginY, this.OneDBeginY);
            yield return new Internal.CellTuple(nameof(SRCCON.OneDEndX), SRCCON.OneDEndX, this.OneDEndX);
            yield return new Internal.CellTuple(nameof(SRCCON.OneDEndY), SRCCON.OneDEndY, this.OneDEndY);

            // Shape Other
            yield return new Internal.CellTuple(nameof(SRCCON.GroupSelectMode), SRCCON.GroupSelectMode, this.GroupSelectMode);

            // Shape Character
            yield return new Internal.CellTuple(nameof(SRCCON.CharCase), SRCCON.CharCase, this.CharCase);
            yield return new Internal.CellTuple(nameof(SRCCON.CharColor), SRCCON.CharColor, this.CharColor);
            yield return new Internal.CellTuple(nameof(SRCCON.CharColorTransparency), SRCCON.CharColorTransparency,
                this.CharColorTransparency);
            yield return new Internal.CellTuple(nameof(SRCCON.CharFont), SRCCON.CharFont, this.CharFont);
            yield return new Internal.CellTuple(nameof(SRCCON.CharFontScale), SRCCON.CharFontScale, this.CharFontScale);
            yield return new Internal.CellTuple(nameof(SRCCON.CharLetterspace), SRCCON.CharLetterspace, this.CharLetterspace);
            yield return new Internal.CellTuple(nameof(SRCCON.CharSize), SRCCON.CharSize, this.CharSize);
            yield return new Internal.CellTuple(nameof(SRCCON.CharStyle), SRCCON.CharStyle, this.CharStyle);
            yield return new Internal.CellTuple(nameof(SRCCON.CharDoubleStrikethrough), SRCCON.CharDoubleStrikethrough,
                this.CharDoubleStrikethrough);
            yield return new Internal.CellTuple(nameof(SRCCON.CharDoubleUnderline), SRCCON.CharDoubleUnderline,
                this.CharDoubleUnderline);
            yield return new Internal.CellTuple(nameof(SRCCON.CharLangID), SRCCON.CharLangID, this.CharLangID);
            yield return new Internal.CellTuple(nameof(SRCCON.CharLocale), SRCCON.CharLocale, this.CharLocale);
            yield return new Internal.CellTuple(nameof(SRCCON.CharOverline), SRCCON.CharOverline, this.CharOverline);
            yield return new Internal.CellTuple(nameof(SRCCON.CharPerpendicular), SRCCON.CharPerpendicular,
                this.CharPerpendicular);
            yield return new Internal.CellTuple(nameof(SRCCON.CharPos), SRCCON.CharPos, this.CharPos);
            yield return new Internal.CellTuple(nameof(SRCCON.CharStrikethru), SRCCON.CharStrikethru, this.CharStrikethru);

            yield return new Internal.CellTuple(nameof(SRCCON.TextXFormAngle), SRCCON.TextXFormAngle, this.TextFormAngle);
            yield return new Internal.CellTuple(nameof(SRCCON.TextXFormHeight), SRCCON.TextXFormHeight, this.TextFormHeight);
            yield return new Internal.CellTuple(nameof(SRCCON.TextXFormLocPinX), SRCCON.TextXFormLocPinX, this.TextFormLocPinX);
            yield return new Internal.CellTuple(nameof(SRCCON.TextXFormLocPinY), SRCCON.TextXFormLocPinY, this.TextFormLocPinY);
            yield return new Internal.CellTuple(nameof(SRCCON.TextXFormPinX), SRCCON.TextXFormPinX, this.TextFormPinX);
            yield return new Internal.CellTuple(nameof(SRCCON.TextXFormPinY), SRCCON.TextXFormPinY, this.TextFormPinY);
            yield return new Internal.CellTuple(nameof(SRCCON.TextXFormWidth), SRCCON.TextXFormWidth, this.TextFormWidth);
        }
    }
}

