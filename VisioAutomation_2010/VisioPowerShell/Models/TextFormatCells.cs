using System.Collections.Generic;
using SRCCON = VisioAutomation.ShapeSheet.SrcConstants;

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
            yield return new CellTuple(nameof(SRCCON.CharCase), SRCCON.CharCase, this.CharCase);
            yield return new CellTuple(nameof(SRCCON.CharColor), SRCCON.CharColor, this.CharColor);
            yield return new CellTuple(nameof(SRCCON.CharColorTransparency), SRCCON.CharColorTransparency, this.CharColorTransparency);
            yield return new CellTuple(nameof(SRCCON.CharFont), SRCCON.CharFont, this.CharFont);
            yield return new CellTuple(nameof(SRCCON.CharFontScale), SRCCON.CharFontScale, this.CharFontScale);
            yield return new CellTuple(nameof(SRCCON.CharLetterspace), SRCCON.CharLetterspace, this.CharLetterspace);
            yield return new CellTuple(nameof(SRCCON.CharSize), SRCCON.CharSize, this.CharSize);

            yield return new CellTuple(nameof(SRCCON.CharDoubleStrikethrough), SRCCON.CharDoubleStrikethrough, this.CharDoubleStrikethrough);
            yield return new CellTuple(nameof(SRCCON.CharDoubleUnderline), SRCCON.CharDoubleUnderline, this.CharDoubleUnderline);
            yield return new CellTuple(nameof(SRCCON.CharLangID), SRCCON.CharLangID, this.CharLangID);
            yield return new CellTuple(nameof(SRCCON.CharLocale), SRCCON.CharLocale, this.CharLocale);
            yield return new CellTuple(nameof(SRCCON.CharOverline), SRCCON.CharOverline, this.CharOverline);
            yield return new CellTuple(nameof(SRCCON.CharPerpendicular), SRCCON.CharPerpendicular, this.CharPerpendicular);
            yield return new CellTuple(nameof(SRCCON.CharPos), SRCCON.CharPos, this.CharPos);
            yield return new CellTuple(nameof(SRCCON.CharStrikethru), SRCCON.CharStrikethru, this.CharStrikethru);
        }
    }
}