using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet.CellGroups.Queries;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Text
{
    public class CharacterCells : ShapeSheet.CellGroups.CellGroupMultiRow
    {
        public ShapeSheet.CellData Color { get; set; }
        public ShapeSheet.CellData Font { get; set; }
        public ShapeSheet.CellData Size { get; set; }
        public ShapeSheet.CellData Style { get; set; }
        public ShapeSheet.CellData Transparency { get; set; }
        public ShapeSheet.CellData AsianFont { get; set; }
        public ShapeSheet.CellData Case { get; set; }
        public ShapeSheet.CellData ComplexScriptFont { get; set; }
        public ShapeSheet.CellData ComplexScriptSize { get; set; }
        public ShapeSheet.CellData DoubleStrikeThrough { get; set; }
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

        public override IEnumerable<SRCFormulaPair> Pairs
        {
            get
            {
                yield return this.newpair(ShapeSheet.SRCConstants.CharColor, this.Color.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.CharFont, this.Font.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.CharSize, this.Size.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.CharStyle, this.Style.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.CharColorTrans, this.Transparency.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.CharAsianFont, this.AsianFont.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.CharCase, this.Case.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.CharComplexScriptFont, this.ComplexScriptFont.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.CharComplexScriptSize, this.ComplexScriptSize.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.CharDblUnderline, this.DoubleUnderline.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.CharDoubleStrikethrough, this.DoubleStrikeThrough.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.CharLangID, this.LangID.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.CharFontScale, this.FontScale.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.CharLangID, this.LangID.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.CharLetterspace, this.Letterspace.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.CharLocale, this.Locale.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.CharLocalizeFont, this.LocalizeFont.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.CharOverline, this.Overline.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.CharPerpendicular, this.Perpendicular.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.CharPos, this.Pos.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.CharRTLText, this.RTLText.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.CharStrikethru, this.Strikethru.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.CharUseVertical, this.UseVertical.Formula);
            }
        }

        public static IList<List<CharacterCells>> GetCells(IVisio.Page page, IList<int> shapeids)
        {
            var query = CharacterCells.lazy_query.Value;
            return query.GetCellGroups(page, shapeids);
        }

        public static IList<CharacterCells> GetCells(IVisio.Shape shape)
        {
            var query = CharacterCells.lazy_query.Value;
            return query.GetCellGroups(shape);
        }

        private static System.Lazy<CharacterFormatCellsQuery> lazy_query = new System.Lazy<CharacterFormatCellsQuery>();


    }
}