using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet.Query;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Text
{
    class CharacterFormatCellsReader : ReaderMultiRow<Text.CharacterFormatCells>
    {
        public SectionSubQueryColumn Font { get; set; }
        public SectionSubQueryColumn Style { get; set; }
        public SectionSubQueryColumn Color { get; set; }
        public SectionSubQueryColumn Size { get; set; }
        public SectionSubQueryColumn Trans { get; set; }
        public SectionSubQueryColumn AsianFont { get; set; }
        public SectionSubQueryColumn Case { get; set; }
        public SectionSubQueryColumn ComplexScriptFont { get; set; }
        public SectionSubQueryColumn ComplexScriptSize { get; set; }
        public SectionSubQueryColumn DoubleStrikethrough { get; set; }
        public SectionSubQueryColumn DoubleUnderline { get; set; }
        public SectionSubQueryColumn LangID { get; set; }
        public SectionSubQueryColumn Locale { get; set; }
        public SectionSubQueryColumn LocalizeFont { get; set; }
        public SectionSubQueryColumn Overline { get; set; }
        public SectionSubQueryColumn Perpendicular { get; set; }
        public SectionSubQueryColumn Pos { get; set; }
        public SectionSubQueryColumn RtlText { get; set; }
        public SectionSubQueryColumn FontScale { get; set; }
        public SectionSubQueryColumn Letterspace { get; set; }
        public SectionSubQueryColumn Strikethru { get; set; }
        public SectionSubQueryColumn UseVertical { get; set; }

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