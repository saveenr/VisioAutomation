using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Text
{
    public class ParagraphCells : ShapeSheetQuery.QueryGroups.CellQueryGroupMultiRow
    {
        public ShapeSheet.CellData<double> IndentFirst { get; set; }
        public ShapeSheet.CellData<double> IndentRight { get; set; }
        public ShapeSheet.CellData<double> IndentLeft { get; set; }
        public ShapeSheet.CellData<double> SpacingBefore { get; set; }
        public ShapeSheet.CellData<double> SpacingAfter { get; set; }
        public ShapeSheet.CellData<double> SpacingLine { get; set; }
        public ShapeSheet.CellData<int> HorizontalAlign { get; set; }
        public ShapeSheet.CellData<int> Bullet { get; set; }
        public ShapeSheet.CellData<int> BulletFont { get; set; }
        public ShapeSheet.CellData<int> BulletFontSize { get; set; }
        public ShapeSheet.CellData<int> LocBulletFont { get; set; }
        public ShapeSheet.CellData<double> TextPosAfterBullet { get; set; }
        public ShapeSheet.CellData<int> Flags { get; set; }
        public ShapeSheet.CellData<string> BulletString { get; set; }

        public override IEnumerable<SRCFormulaPair> Pairs
        {
            get
            {
                yield return this.newpair(ShapeSheet.SRCConstants.Para_IndLeft, this.IndentLeft.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.Para_IndFirst, this.IndentFirst.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.Para_IndRight, this.IndentRight.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.Para_SpAfter, this.SpacingAfter.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.Para_SpBefore, this.SpacingBefore.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.Para_SpLine, this.SpacingLine.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.Para_HorzAlign, this.HorizontalAlign.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.Para_BulletFont, this.BulletFont.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.Para_Bullet, this.Bullet.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.Para_BulletFontSize, this.BulletFontSize.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.Para_LocalizeBulletFont, this.LocBulletFont.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.Para_TextPosAfterBullet, this.TextPosAfterBullet.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.Para_Flags, this.Flags.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.Para_BulletStr, this.BulletString.Formula);
            }
        }

        public static IList<List<ParagraphCells>> GetCells(IVisio.Page page, IList<int> shapeids)
        {
            var query = ParagraphCells.lazy_query.Value;
            return ShapeSheetQuery.QueryGroups.CellQueryGroupMultiRow._GetCells<ParagraphCells, double>(page, shapeids, query, query.GetCells);
        }

        public static IList<ParagraphCells> GetCells(IVisio.Shape shape)
        {
            var query = ParagraphCells.lazy_query.Value;
            return ShapeSheetQuery.QueryGroups.CellQueryGroupMultiRow._GetCells<ParagraphCells, double>(shape, query, query.GetCells);
        }

        private static System.Lazy<ShapeSheetQuery.CommonQueries.ParagraphFormatCellsQuery> lazy_query = new System.Lazy<ShapeSheetQuery.CommonQueries.ParagraphFormatCellsQuery>();
    }
} 