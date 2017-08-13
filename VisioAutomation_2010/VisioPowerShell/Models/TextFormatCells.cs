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

        public string CharDoubleStrikethrough { get; set; }
        public string CharDoubleUnderline { get; set; }
        public string CharLangID { get; set; }
        public string CharLocale { get; set; }
        public string CharOverline { get; set; }
        public string CharPerpendicular { get; set; }
        public string CharPos { get; set; }
        public string CharStrikethru { get; set; }

        public override IEnumerable<CellTuple> GetCellTuples()
        {
            yield return new CellTuple(nameof(SrcConstants.CharCase), SrcConstants.CharCase, this.CharCase);
            yield return new CellTuple(nameof(SrcConstants.CharColor), SrcConstants.CharColor, this.CharColor);
            yield return new CellTuple(nameof(SrcConstants.CharColorTransparency), SrcConstants.CharColorTransparency, this.CharColorTransparency);
            yield return new CellTuple(nameof(SrcConstants.CharFont), SrcConstants.CharFont, this.CharFont);
            yield return new CellTuple(nameof(SrcConstants.CharFontScale), SrcConstants.CharFontScale, this.CharFontScale);
            yield return new CellTuple(nameof(SrcConstants.CharLetterspace), SrcConstants.CharLetterspace, this.CharLetterspace);
            yield return new CellTuple(nameof(SrcConstants.CharSize), SrcConstants.CharSize, this.CharSize);

            yield return new CellTuple(nameof(SrcConstants.CharDoubleStrikethrough), SrcConstants.CharDoubleStrikethrough, this.CharDoubleStrikethrough);
            yield return new CellTuple(nameof(SrcConstants.CharDoubleUnderline), SrcConstants.CharDoubleUnderline, this.CharDoubleUnderline);
            yield return new CellTuple(nameof(SrcConstants.CharLangID), SrcConstants.CharLangID, this.CharLangID);
            yield return new CellTuple(nameof(SrcConstants.CharLocale), SrcConstants.CharLocale, this.CharLocale);
            yield return new CellTuple(nameof(SrcConstants.CharOverline), SrcConstants.CharOverline, this.CharOverline);
            yield return new CellTuple(nameof(SrcConstants.CharPerpendicular), SrcConstants.CharPerpendicular, this.CharPerpendicular);
            yield return new CellTuple(nameof(SrcConstants.CharPos), SrcConstants.CharPos, this.CharPos);
            yield return new CellTuple(nameof(SrcConstants.CharStrikethru), SrcConstants.CharStrikethru, this.CharStrikethru);
        }
    }
}