using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Text
{
    public class ParagraphCells : ShapeSheet.CellGroups.CellGroupMultiRow
    {
        public ShapeSheet.CellData IndentFirst { get; set; }
        public ShapeSheet.CellData IndentRight { get; set; }
        public ShapeSheet.CellData IndentLeft { get; set; }
        public ShapeSheet.CellData SpacingBefore { get; set; }
        public ShapeSheet.CellData SpacingAfter { get; set; }
        public ShapeSheet.CellData SpacingLine { get; set; }
        public ShapeSheet.CellData HorizontalAlign { get; set; }
        public ShapeSheet.CellData Bullet { get; set; }
        public ShapeSheet.CellData BulletFont { get; set; }
        public ShapeSheet.CellData BulletFontSize { get; set; }
        public ShapeSheet.CellData LocBulletFont { get; set; }
        public ShapeSheet.CellData TextPosAfterBullet { get; set; }
        public ShapeSheet.CellData Flags { get; set; }
        public ShapeSheet.CellData BulletString { get; set; }

        public override IEnumerable<SrcFormulaPair> SrcFormulaPairs
        {
            get
            {
                yield return this.newpair(ShapeSheet.SrcConstants.ParaIndentLeft, this.IndentLeft.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.ParaIndentFirst, this.IndentFirst.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.ParaIndentRight, this.IndentRight.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.ParaSpacingAfter, this.SpacingAfter.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.ParaSpacingBefore, this.SpacingBefore.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.ParaSpacingLine, this.SpacingLine.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.ParaHorizontalAlign, this.HorizontalAlign.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.ParaBulletFont, this.BulletFont.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.ParaBullet, this.Bullet.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.ParaBulletFontSize, this.BulletFontSize.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.ParaLocalizeBulletFont, this.LocBulletFont.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.ParaTextPosAfterBullet, this.TextPosAfterBullet.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.ParaFlags, this.Flags.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.ParaBulletStr, this.BulletString.Formula);
            }
        }

        public static List<List<ParagraphCells>> GetCells(IVisio.Page page, IList<int> shapeids)
        {
            var query = ParagraphCells.lazy_query.Value;
            return query.GetCellGroups(page, shapeids);
        }

        public static List<ParagraphCells> GetCells(IVisio.Shape shape)
        {
            var query = ParagraphCells.lazy_query.Value;
            return query.GetCellGroups(shape);
        }

        private static readonly System.Lazy<ParagraphFormatCellsReader> lazy_query = new System.Lazy<ParagraphFormatCellsReader>();
    }
} 