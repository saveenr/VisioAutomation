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
                yield return this.newpair(ShapeSheet.SrcConstants.ParaIndentLeft, this.IndentLeft.ValueF);
                yield return this.newpair(ShapeSheet.SrcConstants.ParaIndentFirst, this.IndentFirst.ValueF);
                yield return this.newpair(ShapeSheet.SrcConstants.ParaIndentRight, this.IndentRight.ValueF);
                yield return this.newpair(ShapeSheet.SrcConstants.ParaSpacingAfter, this.SpacingAfter.ValueF);
                yield return this.newpair(ShapeSheet.SrcConstants.ParaSpacingBefore, this.SpacingBefore.ValueF);
                yield return this.newpair(ShapeSheet.SrcConstants.ParaSpacingLine, this.SpacingLine.ValueF);
                yield return this.newpair(ShapeSheet.SrcConstants.ParaHorizontalAlign, this.HorizontalAlign.ValueF);
                yield return this.newpair(ShapeSheet.SrcConstants.ParaBulletFont, this.BulletFont.ValueF);
                yield return this.newpair(ShapeSheet.SrcConstants.ParaBullet, this.Bullet.ValueF);
                yield return this.newpair(ShapeSheet.SrcConstants.ParaBulletFontSize, this.BulletFontSize.ValueF);
                yield return this.newpair(ShapeSheet.SrcConstants.ParaLocalizeBulletFont, this.LocalizeBulletFont.ValueF);
                yield return this.newpair(ShapeSheet.SrcConstants.ParaTextPosAfterBullet, this.TextPosAfterBullet.ValueF);
                yield return this.newpair(ShapeSheet.SrcConstants.ParaFlags, this.Flags.ValueF);
                yield return this.newpair(ShapeSheet.SrcConstants.ParaBulletString, this.BulletString.ValueF);
            }
        }

        public static List<List<ParagraphFormatCells>> GetCells(IVisio.Page page, IList<int> shapeids)
        {
            var query = ParagraphFormatCells.lazy_query.Value;
            return query.GetCellGroups(page, shapeids);
        }

        public static List<ParagraphFormatCells> GetCells(IVisio.Shape shape)
        {
            var query = ParagraphFormatCells.lazy_query.Value;
            return query.GetCellGroups(shape);
        }

        private static readonly System.Lazy<ParagraphFormatCellsReader> lazy_query = new System.Lazy<ParagraphFormatCellsReader>();
    }
} 