using SRCCON = VisioAutomation.ShapeSheet.SRCConstants;

namespace VisioAutomation.ShapeSheet.Query.Common
{
    class CharacterFormatCellsQuery : CellQuery
    {
        public Query.CellColumn Font { get; set; }
        public Query.CellColumn Style { get; set; }
        public Query.CellColumn Color { get; set; }
        public Query.CellColumn Size { get; set; }
        public Query.CellColumn Trans { get; set; }
        public Query.CellColumn AsianFont { get; set; }
        public Query.CellColumn Case { get; set; }
        public Query.CellColumn ComplexScriptFont { get; set; }
        public Query.CellColumn ComplexScriptSize { get; set; }
        public Query.CellColumn DoubleStrikethrough { get; set; }
        public Query.CellColumn DoubleUnderline { get; set; }
        public Query.CellColumn LangID { get; set; }
        public Query.CellColumn Locale { get; set; }
        public Query.CellColumn LocalizeFont { get; set; }
        public Query.CellColumn Overline { get; set; }
        public Query.CellColumn Perpendicular { get; set; }
        public Query.CellColumn Pos { get; set; }
        public Query.CellColumn RTLText { get; set; }
        public Query.CellColumn FontScale { get; set; }
        public Query.CellColumn Letterspace { get; set; }
        public Query.CellColumn Strikethru { get; set; }
        public Query.CellColumn UseVertical { get; set; }

        public CharacterFormatCellsQuery()
        {
            var sec = this.AddSection(Microsoft.Office.Interop.Visio.VisSectionIndices.visSectionCharacter);

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

        public VisioAutomation.Text.CharacterCells GetCells(System.Collections.Generic.IList<ShapeSheet.CellData<double>> row)
        {
            var cells = new VisioAutomation.Text.CharacterCells();
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