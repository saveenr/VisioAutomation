using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Text
{
    public class ParagraphFormatCells : ShapeSheet.CellGroups.CellGroupMultiRow
    {
        public VisioAutomation.ShapeSheet.CellValueLiteral IndentFirst { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral IndentRight { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral IndentLeft { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral SpacingBefore { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral SpacingAfter { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral SpacingLine { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral HorizontalAlign { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral Bullet { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral BulletFont { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral BulletFontSize { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral LocalizeBulletFont { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral TextPosAfterBullet { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral Flags { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral BulletString { get; set; }

        public override IEnumerable<SrcValuePair> SrcValuePairs
        {
            get
            {
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.ParaIndentLeft, this.IndentLeft.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.ParaIndentFirst, this.IndentFirst.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.ParaIndentRight, this.IndentRight.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.ParaSpacingAfter, this.SpacingAfter.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.ParaSpacingBefore, this.SpacingBefore.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.ParaSpacingLine, this.SpacingLine.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.ParaHorizontalAlign, this.HorizontalAlign.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.ParaBulletFont, this.BulletFont.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.ParaBullet, this.Bullet.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.ParaBulletFontSize, this.BulletFontSize.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.ParaLocalizeBulletFont, this.LocalizeBulletFont.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.ParaTextPosAfterBullet, this.TextPosAfterBullet.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.ParaFlags, this.Flags.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.ParaBulletString, this.BulletString.Value);
            }
        }

        public static List<List<ParagraphFormatCells>> GetFormulas(IVisio.Page page, IList<int> shapeids)
        {
            var query = ParagraphFormatCells.lazy_query.Value;
            return query.GetFormulas(page, shapeids);
        }

        public static List<List<ParagraphFormatCells>> GetResults(IVisio.Page page, IList<int> shapeids)
        {
            var query = ParagraphFormatCells.lazy_query.Value;
            return query.GetResults(page, shapeids);
        }


        public static List<ParagraphFormatCells> GetFormulas(IVisio.Shape shape)
        {
            var query = ParagraphFormatCells.lazy_query.Value;
            return query.GetFormulas(shape);
        }


        public static List<ParagraphFormatCells> GetResults(IVisio.Shape shape)
        {
            var query = ParagraphFormatCells.lazy_query.Value;
            return query.GetResults(shape);
        }

        private static readonly System.Lazy<ParagraphFormatCellsReader> lazy_query = new System.Lazy<ParagraphFormatCellsReader>();
    }
} 