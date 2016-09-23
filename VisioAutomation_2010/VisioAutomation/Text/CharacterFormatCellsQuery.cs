using VisioAutomation.ShapeSheet.CellGroups.Queries;
using VisioAutomation.ShapeSheet.Queries.Columns;
using SRCCON = VisioAutomation.ShapeSheet.SRCConstants;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Text
{
    class CharacterFormatCellsQuery : CellGroupMultiRowQuery<Text.CharacterCells>
    {
        public ColumnSubQuery Font { get; set; }
        public ColumnSubQuery Style { get; set; }
        public ColumnSubQuery Color { get; set; }
        public ColumnSubQuery Size { get; set; }
        public ColumnSubQuery Trans { get; set; }
        public ColumnSubQuery AsianFont { get; set; }
        public ColumnSubQuery Case { get; set; }
        public ColumnSubQuery ComplexScriptFont { get; set; }
        public ColumnSubQuery ComplexScriptSize { get; set; }
        public ColumnSubQuery DoubleStrikethrough { get; set; }
        public ColumnSubQuery DoubleUnderline { get; set; }
        public ColumnSubQuery LangID { get; set; }
        public ColumnSubQuery Locale { get; set; }
        public ColumnSubQuery LocalizeFont { get; set; }
        public ColumnSubQuery Overline { get; set; }
        public ColumnSubQuery Perpendicular { get; set; }
        public ColumnSubQuery Pos { get; set; }
        public ColumnSubQuery RTLText { get; set; }
        public ColumnSubQuery FontScale { get; set; }
        public ColumnSubQuery Letterspace { get; set; }
        public ColumnSubQuery Strikethru { get; set; }
        public ColumnSubQuery UseVertical { get; set; }

        public CharacterFormatCellsQuery()
        {
            var sec = this.query.AddSubQuery(IVisio.VisSectionIndices.visSectionCharacter);

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

        public override Text.CharacterCells CellDataToCellGroup(ShapeSheet.CellData[] row)
        {
            var cells = new Text.CharacterCells();
            cells.Color = row[this.Color];
            cells.Transparency = row[this.Trans];
            cells.Font = row[this.Font];
            cells.Size = row[this.Size];
            cells.Style = row[this.Style];
            cells.AsianFont = row[this.AsianFont];
            cells.AsianFont = row[this.AsianFont];
            cells.Case = row[this.Case];
            cells.ComplexScriptFont = row[this.ComplexScriptFont];
            cells.ComplexScriptSize = row[this.ComplexScriptSize];
            cells.DoubleStrikeThrough = row[this.DoubleStrikethrough];
            cells.DoubleUnderline = row[this.DoubleUnderline];
            cells.FontScale = row[this.FontScale];
            cells.LangID = row[this.LangID];
            cells.Letterspace = row[this.Letterspace];
            cells.Locale = row[this.Locale];
            cells.LocalizeFont = row[this.LocalizeFont];
            cells.Overline = row[this.Overline];
            cells.Perpendicular = row[this.Perpendicular];
            cells.Pos = row[this.Pos];
            cells.RTLText = row[this.RTLText];
            cells.Strikethru = row[this.Strikethru];
            cells.UseVertical = row[this.UseVertical];

            return cells;
        }
    }
}