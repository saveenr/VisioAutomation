using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet.Query;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Text
{
    class CharacterFormatCellsReader : ReaderMultiRow<Text.CharacterFormatCells>
    {
        public SectionQueryColumn Font { get; set; }
        public SectionQueryColumn Style { get; set; }
        public SectionQueryColumn Color { get; set; }
        public SectionQueryColumn Size { get; set; }
        public SectionQueryColumn Trans { get; set; }
        public SectionQueryColumn AsianFont { get; set; }
        public SectionQueryColumn Case { get; set; }
        public SectionQueryColumn ComplexScriptFont { get; set; }
        public SectionQueryColumn ComplexScriptSize { get; set; }
        public SectionQueryColumn DoubleStrikethrough { get; set; }
        public SectionQueryColumn DoubleUnderline { get; set; }
        public SectionQueryColumn LangID { get; set; }
        public SectionQueryColumn Locale { get; set; }
        public SectionQueryColumn LocalizeFont { get; set; }
        public SectionQueryColumn Overline { get; set; }
        public SectionQueryColumn Perpendicular { get; set; }
        public SectionQueryColumn Pos { get; set; }
        public SectionQueryColumn RtlText { get; set; }
        public SectionQueryColumn FontScale { get; set; }
        public SectionQueryColumn Letterspace { get; set; }
        public SectionQueryColumn Strikethru { get; set; }
        public SectionQueryColumn UseVertical { get; set; }

        public CharacterFormatCellsReader()
        {
            var sec = this.query.SectionQueries.Add(IVisio.VisSectionIndices.visSectionCharacter);

            this.Color = sec.Columns.Add(SrcConstants.CharColor, nameof(SrcConstants.CharColor));
            this.Trans = sec.Columns.Add(SrcConstants.CharColorTransparency, nameof(SrcConstants.CharColorTransparency));
            this.Font = sec.Columns.Add(SrcConstants.CharFont, nameof(SrcConstants.CharFont));
            this.Size = sec.Columns.Add(SrcConstants.CharSize, nameof(SrcConstants.CharSize));
            this.Style = sec.Columns.Add(SrcConstants.CharStyle, nameof(SrcConstants.CharStyle));
            this.AsianFont = sec.Columns.Add(SrcConstants.CharAsianFont, nameof(SrcConstants.CharAsianFont));
            this.Case = sec.Columns.Add(SrcConstants.CharCase, nameof(SrcConstants.CharCase));
            this.ComplexScriptFont = sec.Columns.Add(SrcConstants.CharComplexScriptFont, nameof(SrcConstants.CharComplexScriptFont));
            this.ComplexScriptSize = sec.Columns.Add(SrcConstants.CharComplexScriptSize, nameof(SrcConstants.CharComplexScriptSize));
            this.DoubleStrikethrough = sec.Columns.Add(SrcConstants.CharDoubleStrikethrough, nameof(SrcConstants.CharDoubleStrikethrough));
            this.DoubleUnderline = sec.Columns.Add(SrcConstants.CharDoubleUnderline, nameof(SrcConstants.CharDoubleUnderline));
            this.LangID = sec.Columns.Add(SrcConstants.CharLangID, nameof(SrcConstants.CharLangID));
            this.Locale = sec.Columns.Add(SrcConstants.CharLocale, nameof(SrcConstants.CharLocale));
            this.LocalizeFont = sec.Columns.Add(SrcConstants.CharLocalizeFont, nameof(SrcConstants.CharLocalizeFont));
            this.Overline = sec.Columns.Add(SrcConstants.CharOverline, nameof(SrcConstants.CharOverline));
            this.Perpendicular = sec.Columns.Add(SrcConstants.CharPerpendicular, nameof(SrcConstants.CharPerpendicular));
            this.Pos = sec.Columns.Add(SrcConstants.CharPos, nameof(SrcConstants.CharPos));
            this.RtlText = sec.Columns.Add(SrcConstants.CharRTLText, nameof(SrcConstants.CharRTLText));
            this.FontScale = sec.Columns.Add(SrcConstants.CharFontScale, nameof(SrcConstants.CharFontScale));
            this.Letterspace = sec.Columns.Add(SrcConstants.CharLetterspace, nameof(SrcConstants.CharLetterspace));
            this.Strikethru = sec.Columns.Add(SrcConstants.CharStrikethru, nameof(SrcConstants.CharStrikethru));
            this.UseVertical = sec.Columns.Add(SrcConstants.CharUseVertical, nameof(SrcConstants.CharUseVertical));

        }

        public override Text.CharacterFormatCells CellDataToCellGroup(VisioAutomation.Utilities.ArraySegment<string> row)
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