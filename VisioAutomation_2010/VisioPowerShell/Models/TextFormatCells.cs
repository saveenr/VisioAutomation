using System.Collections.Generic;
using VisioAutomation.ShapeSheet;

namespace VisioPowerShell.Models
{
    public class TextFormatCells : VisioPowerShell.Models.BaseCells
    {
        public string CharCase;
        public string CharColor;
        public string CharColorTransparency;
        public string CharFont;
        public string CharFontScale;
        public string CharLetterspace;
        public string CharSize;
        public string CharStyle;

        public string TextXFormAngle;
        public string TextXFormHeight;
        public string TextXFormLocPinX;
        public string TextXFormLocPinY;
        public string TextXFormPinX;
        public string TextXFormPinY;
        public string TextXFormWidth;

        public override IEnumerable<CellTuple> GetCellTuples()
        {
            yield return new CellTuple(nameof(SrcConstants.CharCase), SrcConstants.CharCase, this.CharCase);
            yield return new CellTuple(nameof(SrcConstants.CharColor), SrcConstants.CharColor, this.CharColor);
            yield return new CellTuple(nameof(SrcConstants.CharColorTransparency), SrcConstants.CharColorTransparency, this.CharColorTransparency);
            yield return new CellTuple(nameof(SrcConstants.CharFont), SrcConstants.CharFont, this.CharFont);
            yield return new CellTuple(nameof(SrcConstants.CharFontScale), SrcConstants.CharFontScale, this.CharFontScale);
            yield return new CellTuple(nameof(SrcConstants.CharLetterspace), SrcConstants.CharLetterspace, this.CharLetterspace);
            yield return new CellTuple(nameof(SrcConstants.CharSize), SrcConstants.CharSize, this.CharSize);
            yield return new CellTuple(nameof(SrcConstants.CharStyle), SrcConstants.CharStyle, this.CharStyle);
            yield return new CellTuple(nameof(SrcConstants.TextXFormAngle), SrcConstants.TextXFormAngle, this.TextXFormAngle);
            yield return new CellTuple(nameof(SrcConstants.TextXFormHeight), SrcConstants.TextXFormHeight, this.TextXFormHeight);
            yield return new CellTuple(nameof(SrcConstants.TextXFormLocPinX), SrcConstants.TextXFormLocPinX, this.TextXFormLocPinX);
            yield return new CellTuple(nameof(SrcConstants.TextXFormLocPinY), SrcConstants.TextXFormLocPinY, this.TextXFormLocPinY);
            yield return new CellTuple(nameof(SrcConstants.TextXFormPinX), SrcConstants.TextXFormPinX, this.TextXFormPinX);
            yield return new CellTuple(nameof(SrcConstants.TextXFormPinY), SrcConstants.TextXFormPinY, this.TextXFormPinY);
            yield return new CellTuple(nameof(SrcConstants.TextXFormWidth), SrcConstants.TextXFormWidth, this.TextXFormWidth);
        }
    }
}