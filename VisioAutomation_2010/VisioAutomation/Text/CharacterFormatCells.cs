using System.Collections.Generic;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.CellGroups;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Text
{
    public class CharacterFormatCells : ShapeSheet.CellGroups.CellGroupMultiRow
    {
        public VisioAutomation.ShapeSheet.CellValueLiteral Color { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral Font { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral Size { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral Style { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral ColorTransparency { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral AsianFont { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral Case { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral ComplexScriptFont { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral ComplexScriptSize { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral DoubleStrikethrough { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral DoubleUnderline { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral LangID { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral Locale { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral LocalizeFont { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral Overline { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral Perpendicular { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral Pos { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral RTLText { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral FontScale { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral Letterspace { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral Strikethru { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral UseVertical { get; set; }

        public override IEnumerable<VisioAutomation.ShapeSheet.CellGroups.SrcValuePair> SrcValuePairs
        {
            get
            {
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.CharColor, this.Color);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.CharFont, this.Font);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.CharSize, this.Size);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.CharStyle, this.Style);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.CharColorTransparency, this.ColorTransparency);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.CharAsianFont, this.AsianFont);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.CharCase, this.Case);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.CharComplexScriptFont, this.ComplexScriptFont);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.CharComplexScriptSize, this.ComplexScriptSize);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.CharDoubleUnderline, this.DoubleUnderline);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.CharDoubleStrikethrough, this.DoubleStrikethrough);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.CharLangID, this.LangID);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.CharFontScale, this.FontScale);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.CharLangID, this.LangID);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.CharLetterspace, this.Letterspace);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.CharLocale, this.Locale);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.CharLocalizeFont, this.LocalizeFont);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.CharOverline, this.Overline);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.CharPerpendicular, this.Perpendicular);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.CharPos, this.Pos);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.CharRTLText, this.RTLText);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.CharStrikethru, this.Strikethru);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.CharUseVertical, this.UseVertical);
            }
        }

        public static List<List<CharacterFormatCells>> GetFormulas(IVisio.Page page, IList<int> shapeids)
        {
            var query = CharacterFormatCells.lazy_query.Value;
            return query.GetValues(page, shapeids, CellValueType.Formula);
        }

        public static List<List<CharacterFormatCells>> GetResults(IVisio.Page page, IList<int> shapeids)
        {
            var query = CharacterFormatCells.lazy_query.Value;
            return query.GetValues(page, shapeids, CellValueType.Result);
        }

        public static List<CharacterFormatCells> GetFormulas(IVisio.Shape shape)
        {
            var query = CharacterFormatCells.lazy_query.Value;
            return query.GetValues(shape, CellValueType.Formula);
        }

        public static List<CharacterFormatCells> GetResults(IVisio.Shape shape)
        {
            var query = CharacterFormatCells.lazy_query.Value;
            return query.GetValues(shape, CellValueType.Result);
        }

        private static readonly System.Lazy<CharacterFormatCellsReader> lazy_query = new System.Lazy<CharacterFormatCellsReader>();
    }
}