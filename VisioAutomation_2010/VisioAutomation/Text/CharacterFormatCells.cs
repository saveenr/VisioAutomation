using System.Collections.Generic;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.CellGroups;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.ShapeSheet.Query;

namespace VisioAutomation.Text
{
    public class CharacterFormatCells : CellGroupMultiRow
    {
        public CellValueLiteral Color { get; set; }
        public CellValueLiteral Font { get; set; }
        public CellValueLiteral Size { get; set; }
        public CellValueLiteral Style { get; set; }
        public CellValueLiteral ColorTransparency { get; set; }
        public CellValueLiteral AsianFont { get; set; }
        public CellValueLiteral Case { get; set; }
        public CellValueLiteral ComplexScriptFont { get; set; }
        public CellValueLiteral ComplexScriptSize { get; set; }
        public CellValueLiteral DoubleStrikethrough { get; set; }
        public CellValueLiteral DoubleUnderline { get; set; }
        public CellValueLiteral LangID { get; set; }
        public CellValueLiteral Locale { get; set; }
        public CellValueLiteral LocalizeFont { get; set; }
        public CellValueLiteral Overline { get; set; }
        public CellValueLiteral Perpendicular { get; set; }
        public CellValueLiteral Pos { get; set; }
        public CellValueLiteral RTLText { get; set; }
        public CellValueLiteral FontScale { get; set; }
        public CellValueLiteral Letterspace { get; set; }
        public CellValueLiteral Strikethru { get; set; }
        public CellValueLiteral UseVertical { get; set; }

        public override IEnumerable<SrcValuePair> SrcValuePairs
        {
            get
            {
                yield return SrcValuePair.Create(SrcConstants.CharColor, this.Color);
                yield return SrcValuePair.Create(SrcConstants.CharFont, this.Font);
                yield return SrcValuePair.Create(SrcConstants.CharSize, this.Size);
                yield return SrcValuePair.Create(SrcConstants.CharStyle, this.Style);
                yield return SrcValuePair.Create(SrcConstants.CharColorTransparency, this.ColorTransparency);
                yield return SrcValuePair.Create(SrcConstants.CharAsianFont, this.AsianFont);
                yield return SrcValuePair.Create(SrcConstants.CharCase, this.Case);
                yield return SrcValuePair.Create(SrcConstants.CharComplexScriptFont, this.ComplexScriptFont);
                yield return SrcValuePair.Create(SrcConstants.CharComplexScriptSize, this.ComplexScriptSize);
                yield return SrcValuePair.Create(SrcConstants.CharDoubleUnderline, this.DoubleUnderline);
                yield return SrcValuePair.Create(SrcConstants.CharDoubleStrikethrough, this.DoubleStrikethrough);
                yield return SrcValuePair.Create(SrcConstants.CharLangID, this.LangID);
                yield return SrcValuePair.Create(SrcConstants.CharFontScale, this.FontScale);
                yield return SrcValuePair.Create(SrcConstants.CharLangID, this.LangID);
                yield return SrcValuePair.Create(SrcConstants.CharLetterspace, this.Letterspace);
                yield return SrcValuePair.Create(SrcConstants.CharLocale, this.Locale);
                yield return SrcValuePair.Create(SrcConstants.CharLocalizeFont, this.LocalizeFont);
                yield return SrcValuePair.Create(SrcConstants.CharOverline, this.Overline);
                yield return SrcValuePair.Create(SrcConstants.CharPerpendicular, this.Perpendicular);
                yield return SrcValuePair.Create(SrcConstants.CharPos, this.Pos);
                yield return SrcValuePair.Create(SrcConstants.CharRTLText, this.RTLText);
                yield return SrcValuePair.Create(SrcConstants.CharStrikethru, this.Strikethru);
                yield return SrcValuePair.Create(SrcConstants.CharUseVertical, this.UseVertical);
            }
        }

        public static List<List<CharacterFormatCells>> GetValues(IVisio.Page page, IList<int> shapeids, CellValueType cvt)
        {
            var query = lazy_query.Value;
            return query.GetValues(page, shapeids, cvt);
        }

        public static List<CharacterFormatCells> GetValues(IVisio.Shape shape, CellValueType cvt)
        {
            var query = lazy_query.Value;
            return query.GetValues(shape, cvt);
        }

        private static readonly System.Lazy<CharacterFormatCellsReader> lazy_query = new System.Lazy<CharacterFormatCellsReader>();


        class CharacterFormatCellsReader : ReaderMultiRow<Text.CharacterFormatCells>
        {
            public SectionQueryColumn Font { get; set; }
            public SectionQueryColumn Style { get; set; }
            public SectionQueryColumn Color { get; set; }
            public SectionQueryColumn Size { get; set; }
            public SectionQueryColumn ColorTransparency { get; set; }
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
                this.ColorTransparency = sec.Columns.Add(SrcConstants.CharColorTransparency, nameof(this.ColorTransparency));
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
                cells.ColorTransparency = row[this.ColorTransparency];
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
}