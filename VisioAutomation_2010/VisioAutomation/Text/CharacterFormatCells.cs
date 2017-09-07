using System.Collections.Generic;
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

        public override IEnumerable<VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair> SrcFormulaPairs
        {
            get
            {
                yield return this.newpair(ShapeSheet.SrcConstants.CharColor, this.Color.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.CharFont, this.Font.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.CharSize, this.Size.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.CharStyle, this.Style.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.CharColorTransparency, this.ColorTransparency.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.CharAsianFont, this.AsianFont.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.CharCase, this.Case.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.CharComplexScriptFont, this.ComplexScriptFont.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.CharComplexScriptSize, this.ComplexScriptSize.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.CharDoubleUnderline, this.DoubleUnderline.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.CharDoubleStrikethrough, this.DoubleStrikethrough.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.CharLangID, this.LangID.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.CharFontScale, this.FontScale.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.CharLangID, this.LangID.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.CharLetterspace, this.Letterspace.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.CharLocale, this.Locale.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.CharLocalizeFont, this.LocalizeFont.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.CharOverline, this.Overline.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.CharPerpendicular, this.Perpendicular.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.CharPos, this.Pos.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.CharRTLText, this.RTLText.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.CharStrikethru, this.Strikethru.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.CharUseVertical, this.UseVertical.Value);
            }
        }

        public static List<List<CharacterFormatCells>> GetFormulas(IVisio.Page page, IList<int> shapeids)
        {
            var query = CharacterFormatCells.lazy_query.Value;
            return query.GetFormulas(page, shapeids);
        }

        public static List<List<CharacterFormatCells>> GetResults(IVisio.Page page, IList<int> shapeids)
        {
            var query = CharacterFormatCells.lazy_query.Value;
            return query.GetResults(page, shapeids);
        }

        public static List<CharacterFormatCells> GetFormulas(IVisio.Shape shape)
        {
            var query = CharacterFormatCells.lazy_query.Value;
            return query.GetFormulas(shape);
        }

        public static List<CharacterFormatCells> GetResults(IVisio.Shape shape)
        {
            var query = CharacterFormatCells.lazy_query.Value;
            return query.GetResults(shape);
        }

        private static readonly System.Lazy<CharacterFormatCellsReader> lazy_query = new System.Lazy<CharacterFormatCellsReader>();
    }
}