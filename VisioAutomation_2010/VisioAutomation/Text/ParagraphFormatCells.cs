using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Text
{
    public class ParagraphFormatCells : ShapeSheet.CellGroups.CellGroupMultiRow
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
        public ShapeSheet.CellData LocalizeBulletFont { get; set; }
        public ShapeSheet.CellData TextPosAfterBullet { get; set; }
        public ShapeSheet.CellData Flags { get; set; }
        public ShapeSheet.CellData BulletString { get; set; }

        public override IEnumerable<SrcFormulaPair> SrcFormulaPairs
        {
            get
            {
                yield return this.newpair(ShapeSheet.SrcConstants.ParaIndentLeft, this.IndentLeft.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.ParaIndentFirst, this.IndentFirst.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.ParaIndentRight, this.IndentRight.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.ParaSpacingAfter, this.SpacingAfter.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.ParaSpacingBefore, this.SpacingBefore.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.ParaSpacingLine, this.SpacingLine.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.ParaHorizontalAlign, this.HorizontalAlign.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.ParaBulletFont, this.BulletFont.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.ParaBullet, this.Bullet.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.ParaBulletFontSize, this.BulletFontSize.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.ParaLocalizeBulletFont, this.LocalizeBulletFont.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.ParaTextPosAfterBullet, this.TextPosAfterBullet.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.ParaFlags, this.Flags.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.ParaBulletString, this.BulletString.Value);
            }
        }

        public static List<List<ParagraphFormatCells>> GetCells(IVisio.Page page, IList<int> shapeids, VisioAutomation.ShapeSheet.CellValueType cvt)
        {
            var query = ParagraphFormatCells.lazy_query.Value;
            return query.GetCellGroups(page, shapeids, cvt);
        }

        public static List<ParagraphFormatCells> GetCells(IVisio.Shape shape, VisioAutomation.ShapeSheet.CellValueType cvt)
        {
            var query = ParagraphFormatCells.lazy_query.Value;
            return query.GetCellGroups(shape, cvt);
        }

        private static readonly System.Lazy<ParagraphFormatCellsReader> lazy_query = new System.Lazy<ParagraphFormatCellsReader>();
    }
} 