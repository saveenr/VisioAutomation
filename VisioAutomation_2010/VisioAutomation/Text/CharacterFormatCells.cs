using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Text
{
    public class CharacterFormatCells : ShapeSheet.CellGroups.CellGroupMultiRow
    {
        public ShapeSheet.CellData Color { get; set; }
        public ShapeSheet.CellData Font { get; set; }
        public ShapeSheet.CellData Size { get; set; }
        public ShapeSheet.CellData Style { get; set; }
        public ShapeSheet.CellData ColorTransparency { get; set; }
        public ShapeSheet.CellData AsianFont { get; set; }
        public ShapeSheet.CellData Case { get; set; }
        public ShapeSheet.CellData ComplexScriptFont { get; set; }
        public ShapeSheet.CellData ComplexScriptSize { get; set; }
        public ShapeSheet.CellData DoubleStrikethrough { get; set; }
        public ShapeSheet.CellData DoubleUnderline { get; set; }
        public ShapeSheet.CellData LangID { get; set; }
        public ShapeSheet.CellData Locale { get; set; }
        public ShapeSheet.CellData LocalizeFont { get; set; }
        public ShapeSheet.CellData Overline { get; set; }
        public ShapeSheet.CellData Perpendicular { get; set; }
        public ShapeSheet.CellData Pos { get; set; }
        public ShapeSheet.CellData RTLText { get; set; }
        public ShapeSheet.CellData FontScale { get; set; }
        public ShapeSheet.CellData Letterspace { get; set; }
        public ShapeSheet.CellData Strikethru { get; set; }
        public ShapeSheet.CellData UseVertical { get; set; }

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

        public static List<List<CharacterFormatCells>> GetCells(IVisio.Page page, IList<int> shapeids, VisioAutomation.ShapeSheet.CellValueType cvt)
        {
            var query = CharacterFormatCells.lazy_query.Value;
            return query.GetCellGroups(page, shapeids, cvt);
        }

        public static List<CharacterFormatCells> GetCells(IVisio.Shape shape, VisioAutomation.ShapeSheet.CellValueType cvt)
        {
            var query = CharacterFormatCells.lazy_query.Value;
            return query.GetCellGroups(shape,cvt);
        }

        private static readonly System.Lazy<CharacterFormatCellsReader> lazy_query = new System.Lazy<CharacterFormatCellsReader>();
    }
}