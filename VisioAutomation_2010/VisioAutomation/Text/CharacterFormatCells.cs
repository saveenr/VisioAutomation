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
                yield return this.newpair(ShapeSheet.SrcConstants.CharColor, this.Color.ValueF);
                yield return this.newpair(ShapeSheet.SrcConstants.CharFont, this.Font.ValueF);
                yield return this.newpair(ShapeSheet.SrcConstants.CharSize, this.Size.ValueF);
                yield return this.newpair(ShapeSheet.SrcConstants.CharStyle, this.Style.ValueF);
                yield return this.newpair(ShapeSheet.SrcConstants.CharColorTransparency, this.ColorTransparency.ValueF);
                yield return this.newpair(ShapeSheet.SrcConstants.CharAsianFont, this.AsianFont.ValueF);
                yield return this.newpair(ShapeSheet.SrcConstants.CharCase, this.Case.ValueF);
                yield return this.newpair(ShapeSheet.SrcConstants.CharComplexScriptFont, this.ComplexScriptFont.ValueF);
                yield return this.newpair(ShapeSheet.SrcConstants.CharComplexScriptSize, this.ComplexScriptSize.ValueF);
                yield return this.newpair(ShapeSheet.SrcConstants.CharDoubleUnderline, this.DoubleUnderline.ValueF);
                yield return this.newpair(ShapeSheet.SrcConstants.CharDoubleStrikethrough, this.DoubleStrikethrough.ValueF);
                yield return this.newpair(ShapeSheet.SrcConstants.CharLangID, this.LangID.ValueF);
                yield return this.newpair(ShapeSheet.SrcConstants.CharFontScale, this.FontScale.ValueF);
                yield return this.newpair(ShapeSheet.SrcConstants.CharLangID, this.LangID.ValueF);
                yield return this.newpair(ShapeSheet.SrcConstants.CharLetterspace, this.Letterspace.ValueF);
                yield return this.newpair(ShapeSheet.SrcConstants.CharLocale, this.Locale.ValueF);
                yield return this.newpair(ShapeSheet.SrcConstants.CharLocalizeFont, this.LocalizeFont.ValueF);
                yield return this.newpair(ShapeSheet.SrcConstants.CharOverline, this.Overline.ValueF);
                yield return this.newpair(ShapeSheet.SrcConstants.CharPerpendicular, this.Perpendicular.ValueF);
                yield return this.newpair(ShapeSheet.SrcConstants.CharPos, this.Pos.ValueF);
                yield return this.newpair(ShapeSheet.SrcConstants.CharRTLText, this.RTLText.ValueF);
                yield return this.newpair(ShapeSheet.SrcConstants.CharStrikethru, this.Strikethru.ValueF);
                yield return this.newpair(ShapeSheet.SrcConstants.CharUseVertical, this.UseVertical.ValueF);
            }
        }

        public static List<List<CharacterFormatCells>> GetCells(IVisio.Page page, IList<int> shapeids)
        {
            var query = CharacterFormatCells.lazy_query.Value;
            return query.GetCellGroups(page, shapeids);
        }

        public static List<CharacterFormatCells> GetCells(IVisio.Shape shape)
        {
            var query = CharacterFormatCells.lazy_query.Value;
            return query.GetCellGroups(shape);
        }

        private static readonly System.Lazy<CharacterFormatCellsReader> lazy_query = new System.Lazy<CharacterFormatCellsReader>();
    }
}