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

            this.Color = sec.Columns.Add(SrcConstants.CharColor, nameof(this.Color));
            this.Trans = sec.Columns.Add(SrcConstants.CharColorTransparency, nameof(this.Trans));
            this.Font = sec.Columns.Add(SrcConstants.CharFont, nameof(this.Font));
            this.Size = sec.Columns.Add(SrcConstants.CharSize, nameof(this.Size));
            this.Style = sec.Columns.Add(SrcConstants.CharStyle, nameof(this.Style));
            this.AsianFont = sec.Columns.Add(SrcConstants.CharAsianFont, nameof(this.AsianFont));
            this.Case = sec.Columns.Add(SrcConstants.CharCase, nameof(this.Case));
            this.ComplexScriptFont = sec.Columns.Add(SrcConstants.CharComplexScriptFont, nameof(this.ComplexScriptFont));
            this.ComplexScriptSize = sec.Columns.Add(SrcConstants.CharComplexScriptSize, nameof(this.ComplexScriptSize));
            this.DoubleStrikethrough = sec.Columns.Add(SrcConstants.CharDoubleStrikethrough, nameof(this.DoubleStrikethrough));
            this.DoubleUnderline = sec.Columns.Add(SrcConstants.CharDoubleUnderline, nameof(this.DoubleUnderline));
            this.LangID = sec.Columns.Add(SrcConstants.CharLangID, nameof(this.LangID));
            this.Locale = sec.Columns.Add(SrcConstants.CharLocale, nameof(this.Locale));
            this.LocalizeFont = sec.Columns.Add(SrcConstants.CharLocalizeFont, nameof(this.LocalizeFont));
            this.Overline = sec.Columns.Add(SrcConstants.CharOverline, nameof(this.Overline));
            this.Perpendicular = sec.Columns.Add(SrcConstants.CharPerpendicular, nameof(this.Perpendicular));
            this.Pos = sec.Columns.Add(SrcConstants.CharPos, nameof(this.Pos));
            this.RtlText = sec.Columns.Add(SrcConstants.CharRTLText, nameof(this.RtlText));
            this.FontScale = sec.Columns.Add(SrcConstants.CharFontScale, nameof(this.FontScale));
            this.Letterspace = sec.Columns.Add(SrcConstants.CharLetterspace, nameof(this.Letterspace));
            this.Strikethru = sec.Columns.Add(SrcConstants.CharStrikethru, nameof(this.Strikethru));
            this.UseVertical = sec.Columns.Add(SrcConstants.CharUseVertical, nameof(this.UseVertical));

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