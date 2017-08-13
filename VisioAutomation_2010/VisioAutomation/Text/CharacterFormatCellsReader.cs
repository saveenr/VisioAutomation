using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet.Query;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Text
{
    class CharacterFormatCellsReader : MultiRowReader<Text.CharacterFormatCells>
    {
        public SubQueryColumn Font { get; set; }
        public SubQueryColumn Style { get; set; }
        public SubQueryColumn Color { get; set; }
        public SubQueryColumn Size { get; set; }
        public SubQueryColumn Trans { get; set; }
        public SubQueryColumn AsianFont { get; set; }
        public SubQueryColumn Case { get; set; }
        public SubQueryColumn ComplexScriptFont { get; set; }
        public SubQueryColumn ComplexScriptSize { get; set; }
        public SubQueryColumn DoubleStrikethrough { get; set; }
        public SubQueryColumn DoubleUnderline { get; set; }
        public SubQueryColumn LangID { get; set; }
        public SubQueryColumn Locale { get; set; }
        public SubQueryColumn LocalizeFont { get; set; }
        public SubQueryColumn Overline { get; set; }
        public SubQueryColumn Perpendicular { get; set; }
        public SubQueryColumn Pos { get; set; }
        public SubQueryColumn RtlText { get; set; }
        public SubQueryColumn FontScale { get; set; }
        public SubQueryColumn Letterspace { get; set; }
        public SubQueryColumn Strikethru { get; set; }
        public SubQueryColumn UseVertical { get; set; }

        public CharacterFormatCellsReader()
        {
            var sec = this.query.AddSubQuery(IVisio.VisSectionIndices.visSectionCharacter);

            this.Color = sec.AddCell(SrcConstants.CharColor, nameof(SrcConstants.CharColor));
            this.Trans = sec.AddCell(SrcConstants.CharColorTransparency, nameof(SrcConstants.CharColorTransparency));
            this.Font = sec.AddCell(SrcConstants.CharFont, nameof(SrcConstants.CharFont));
            this.Size = sec.AddCell(SrcConstants.CharSize, nameof(SrcConstants.CharSize));
            this.Style = sec.AddCell(SrcConstants.CharStyle, nameof(SrcConstants.CharStyle));
            this.AsianFont = sec.AddCell(SrcConstants.CharAsianFont, nameof(SrcConstants.CharAsianFont));
            this.Case = sec.AddCell(SrcConstants.CharCase, nameof(SrcConstants.CharCase));
            this.ComplexScriptFont = sec.AddCell(SrcConstants.CharComplexScriptFont, nameof(SrcConstants.CharComplexScriptFont));
            this.ComplexScriptSize = sec.AddCell(SrcConstants.CharComplexScriptSize, nameof(SrcConstants.CharComplexScriptSize));
            this.DoubleStrikethrough = sec.AddCell(SrcConstants.CharDoubleStrikethrough, nameof(SrcConstants.CharDoubleStrikethrough));
            this.DoubleUnderline = sec.AddCell(SrcConstants.CharDoubleUnderline, nameof(SrcConstants.CharDoubleUnderline));
            this.LangID = sec.AddCell(SrcConstants.CharLangID, nameof(SrcConstants.CharLangID));
            this.Locale = sec.AddCell(SrcConstants.CharLocale, nameof(SrcConstants.CharLocale));
            this.LocalizeFont = sec.AddCell(SrcConstants.CharLocalizeFont, nameof(SrcConstants.CharLocalizeFont));
            this.Overline = sec.AddCell(SrcConstants.CharOverline, nameof(SrcConstants.CharOverline));
            this.Perpendicular = sec.AddCell(SrcConstants.CharPerpendicular, nameof(SrcConstants.CharPerpendicular));
            this.Pos = sec.AddCell(SrcConstants.CharPos, nameof(SrcConstants.CharPos));
            this.RtlText = sec.AddCell(SrcConstants.CharRTLText, nameof(SrcConstants.CharRTLText));
            this.FontScale = sec.AddCell(SrcConstants.CharFontScale, nameof(SrcConstants.CharFontScale));
            this.Letterspace = sec.AddCell(SrcConstants.CharLetterspace, nameof(SrcConstants.CharLetterspace));
            this.Strikethru = sec.AddCell(SrcConstants.CharStrikethru, nameof(SrcConstants.CharStrikethru));
            this.UseVertical = sec.AddCell(SrcConstants.CharUseVertical, nameof(SrcConstants.CharUseVertical));

        }

        public override Text.CharacterFormatCells CellDataToCellGroup(VisioAutomation.Utilities.ArraySegment<ShapeSheet.CellData> row)
        {
            var cells = new Text.CharacterFormatCells();
            cells.Color = row[this.Color];
            cells.ColorTransparency = row[this.Trans];
            cells.Font = row[this.Font];
            cells.Size = row[this.Size];
            cells.Style = row[this.Style];
            cells.AsianFont = row[this.AsianFont];
            cells.AsianFont = row[this.AsianFont];
            cells.Case = row[this.Case];
            cells.ComplexScriptFont = row[this.ComplexScriptFont];
            cells.ComplexScriptSize = row[this.ComplexScriptSize];
            cells.DoubleStrikethrough = row[this.DoubleStrikethrough];
            cells.DoubleUnderline = row[this.DoubleUnderline];
            cells.FontScale = row[this.FontScale];
            cells.LangID = row[this.LangID];
            cells.Letterspace = row[this.Letterspace];
            cells.Locale = row[this.Locale];
            cells.LocalizeFont = row[this.LocalizeFont];
            cells.Overline = row[this.Overline];
            cells.Perpendicular = row[this.Perpendicular];
            cells.Pos = row[this.Pos];
            cells.RTLText = row[this.RtlText];
            cells.Strikethru = row[this.Strikethru];
            cells.UseVertical = row[this.UseVertical];

            return cells;
        }
    }
}