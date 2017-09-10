using System.Collections.Generic;
using VisioAutomation.ShapeSheet;
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
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.ParaIndentLeft, this.IndentLeft);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.ParaIndentFirst, this.IndentFirst);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.ParaIndentRight, this.IndentRight);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.ParaSpacingAfter, this.SpacingAfter);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.ParaSpacingBefore, this.SpacingBefore);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.ParaSpacingLine, this.SpacingLine);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.ParaHorizontalAlign, this.HorizontalAlign);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.ParaBulletFont, this.BulletFont);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.ParaBullet, this.Bullet);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.ParaBulletFontSize, this.BulletFontSize);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.ParaLocalizeBulletFont, this.LocalizeBulletFont);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.ParaTextPosAfterBullet, this.TextPosAfterBullet);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.ParaFlags, this.Flags);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.ParaBulletString, this.BulletString);
            }
        }

        public static List<List<ParagraphFormatCells>> GetFormulas(IVisio.Page page, IList<int> shapeids)
        {
            var query = ParagraphFormatCells.lazy_query.Value;
            return query.GetValues(page, shapeids, CellValueType.Formula);
        }

        public static List<List<ParagraphFormatCells>> GetResults(IVisio.Page page, IList<int> shapeids)
        {
            var query = ParagraphFormatCells.lazy_query.Value;
            return query.GetValues(page, shapeids, CellValueType.Result);
        }
        
        public static List<ParagraphFormatCells> GetFormulas(IVisio.Shape shape)
        {
            var query = ParagraphFormatCells.lazy_query.Value;
            return query.GetValues(shape, CellValueType.Formula);
        }


        public static List<ParagraphFormatCells> GetResults(IVisio.Shape shape)
        {
            var query = ParagraphFormatCells.lazy_query.Value;
            return query.GetValues(shape, CellValueType.Result);
        }

        private static readonly System.Lazy<ParagraphFormatCellsReader> lazy_query = new System.Lazy<ParagraphFormatCellsReader>();
    }
} 