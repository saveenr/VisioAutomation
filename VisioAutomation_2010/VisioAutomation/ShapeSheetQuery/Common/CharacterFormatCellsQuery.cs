using SRCCON = VisioAutomation.ShapeSheet.SRCConstants;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheetQuery.Common
{
    class CharacterFormatCellsQuery : CellQuery
    {
        public CellColumn Font { get; set; }
        public CellColumn Style { get; set; }
        public CellColumn Color { get; set; }
        public CellColumn Size { get; set; }
        public CellColumn Trans { get; set; }
        public CellColumn AsianFont { get; set; }
        public CellColumn Case { get; set; }
        public CellColumn ComplexScriptFont { get; set; }
        public CellColumn ComplexScriptSize { get; set; }
        public CellColumn DoubleStrikethrough { get; set; }
        public CellColumn DoubleUnderline { get; set; }
        public CellColumn LangID { get; set; }
        public CellColumn Locale { get; set; }
        public CellColumn LocalizeFont { get; set; }
        public CellColumn Overline { get; set; }
        public CellColumn Perpendicular { get; set; }
        public CellColumn Pos { get; set; }
        public CellColumn RTLText { get; set; }
        public CellColumn FontScale { get; set; }
        public CellColumn Letterspace { get; set; }
        public CellColumn Strikethru { get; set; }
        public CellColumn UseVertical { get; set; }

        public CharacterFormatCellsQuery()
        {
            var sec = this.AddSection(IVisio.VisSectionIndices.visSectionCharacter);

            this.Color = sec.AddCell(SRCCON.CharColor, nameof(SRCCON.CharColor));
            this.Trans = sec.AddCell(SRCCON.CharColorTrans, nameof(SRCCON.CharColorTrans));
            this.Font = sec.AddCell(SRCCON.CharFont, nameof(SRCCON.CharFont));
            this.Size = sec.AddCell(SRCCON.CharSize, nameof(SRCCON.CharSize));
            this.Style = sec.AddCell(SRCCON.CharStyle, nameof(SRCCON.CharStyle));
            this.AsianFont = sec.AddCell(SRCCON.CharAsianFont, nameof(SRCCON.CharAsianFont));
            this.Case = sec.AddCell(SRCCON.CharCase, nameof(SRCCON.CharCase));
            this.ComplexScriptFont = sec.AddCell(SRCCON.CharComplexScriptFont, nameof(SRCCON.CharComplexScriptFont));
            this.ComplexScriptSize = sec.AddCell(SRCCON.CharComplexScriptSize, nameof(SRCCON.CharComplexScriptSize));
            this.DoubleStrikethrough = sec.AddCell(SRCCON.CharDoubleStrikethrough, nameof(SRCCON.CharDoubleStrikethrough));
            this.DoubleUnderline = sec.AddCell(SRCCON.CharDblUnderline, nameof(SRCCON.CharDblUnderline));
            this.LangID = sec.AddCell(SRCCON.CharLangID, nameof(SRCCON.CharLangID));
            this.Locale = sec.AddCell(SRCCON.CharLocale, nameof(SRCCON.CharLocale));
            this.LocalizeFont = sec.AddCell(SRCCON.CharLocalizeFont, nameof(SRCCON.CharLocalizeFont));
            this.Overline = sec.AddCell(SRCCON.CharOverline, nameof(SRCCON.CharOverline));
            this.Perpendicular = sec.AddCell(SRCCON.CharPerpendicular, nameof(SRCCON.CharPerpendicular));
            this.Pos = sec.AddCell(SRCCON.CharPos, nameof(SRCCON.CharPos));
            this.RTLText = sec.AddCell(SRCCON.CharRTLText, nameof(SRCCON.CharRTLText));
            this.FontScale = sec.AddCell(SRCCON.CharFontScale, nameof(SRCCON.CharFontScale));
            this.Letterspace = sec.AddCell(SRCCON.CharLetterspace, nameof(SRCCON.CharLetterspace));
            this.Strikethru = sec.AddCell(SRCCON.CharStrikethru, nameof(SRCCON.CharStrikethru));
            this.UseVertical = sec.AddCell(SRCCON.CharUseVertical, nameof(SRCCON.CharUseVertical));

        }

        public Text.CharacterCells GetCells(ShapeSheet.CellData<double>[] row)
        {
            var cells = new Text.CharacterCells();
            cells.Color = Extensions.CellDataMethods.ToInt(row[this.Color]);
            cells.Transparency = row[this.Trans];
            cells.Font = Extensions.CellDataMethods.ToInt(row[this.Font]);
            cells.Size = row[this.Size];
            cells.Style = Extensions.CellDataMethods.ToInt(row[this.Style]);
            cells.AsianFont = Extensions.CellDataMethods.ToInt(row[this.AsianFont]);
            cells.AsianFont = Extensions.CellDataMethods.ToInt(row[this.AsianFont]);
            cells.Case = Extensions.CellDataMethods.ToInt(row[this.Case]);
            cells.ComplexScriptFont = Extensions.CellDataMethods.ToInt(row[this.ComplexScriptFont]);
            cells.ComplexScriptSize = row[this.ComplexScriptSize];
            cells.DoubleStrikeThrough = Extensions.CellDataMethods.ToBool(row[this.DoubleStrikethrough]);
            cells.DoubleUnderline = Extensions.CellDataMethods.ToBool(row[this.DoubleUnderline]);
            cells.FontScale = row[this.FontScale];
            cells.LangID = Extensions.CellDataMethods.ToInt(row[this.LangID]);
            cells.Letterspace = row[this.Letterspace];
            cells.Locale = Extensions.CellDataMethods.ToInt(row[this.Locale]);
            cells.LocalizeFont = Extensions.CellDataMethods.ToInt(row[this.LocalizeFont]);
            cells.Overline = Extensions.CellDataMethods.ToBool(row[this.Overline]);
            cells.Perpendicular = Extensions.CellDataMethods.ToBool(row[this.Perpendicular]);
            cells.Pos = Extensions.CellDataMethods.ToInt(row[this.Pos]);
            cells.RTLText = Extensions.CellDataMethods.ToInt(row[this.RTLText]);
            cells.Strikethru = Extensions.CellDataMethods.ToBool(row[this.Strikethru]);
            cells.UseVertical = Extensions.CellDataMethods.ToInt(row[this.UseVertical]);

            return cells;
        }
    }
}