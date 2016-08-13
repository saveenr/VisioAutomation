using System.Collections.Generic;
using VisioAutomation.ShapeSheetQuery.QueryGroups;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Text
{
    public class CharacterCells : ShapeSheetQuery.QueryGroups.QueryGroupMultiRow
    {
        public ShapeSheet.CellData<int> Color { get; set; }
        public ShapeSheet.CellData<int> Font { get; set; }
        public ShapeSheet.CellData<double> Size { get; set; }
        public ShapeSheet.CellData<int> Style { get; set; }
        public ShapeSheet.CellData<double> Transparency { get; set; }
        public ShapeSheet.CellData<int> AsianFont { get; set; }
        public ShapeSheet.CellData<int> Case { get; set; }
        public ShapeSheet.CellData<int> ComplexScriptFont { get; set; }
        public ShapeSheet.CellData<double> ComplexScriptSize { get; set; }
        public ShapeSheet.CellData<bool> DoubleStrikeThrough { get; set; }
        public ShapeSheet.CellData<bool> DoubleUnderline { get; set; }
        public ShapeSheet.CellData<int> LangID { get; set; }
        public ShapeSheet.CellData<int> Locale { get; set; }
        public ShapeSheet.CellData<int> LocalizeFont { get; set; }
        public ShapeSheet.CellData<bool> Overline { get; set; }
        public ShapeSheet.CellData<bool> Perpendicular { get; set; }
        public ShapeSheet.CellData<int> Pos { get; set; }
        public ShapeSheet.CellData<int> RTLText { get; set; }
        public ShapeSheet.CellData<double> FontScale { get; set; }
        public ShapeSheet.CellData<double> Letterspace { get; set; }
        public ShapeSheet.CellData<bool> Strikethru { get; set; }
        public ShapeSheet.CellData<int> UseVertical { get; set; }

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
            return ShapeSheetQuery.QueryGroups.QueryGroupMultiRow._GetCells<CharacterCells, double>(page, shapeids, query, query.GetCells);
        }

        public static IList<CharacterCells> GetCells(IVisio.Shape shape)
        {
            var query = CharacterCells.lazy_query.Value;
            return ShapeSheetQuery.QueryGroups.QueryGroupMultiRow._GetCells<CharacterCells, double>(shape, query, query.GetCells);
        }

        private static System.Lazy<ShapeSheetQuery.CommonQueries.CharacterFormatCellsQuery> lazy_query = new System.Lazy<ShapeSheetQuery.CommonQueries.CharacterFormatCellsQuery>();


    }
}