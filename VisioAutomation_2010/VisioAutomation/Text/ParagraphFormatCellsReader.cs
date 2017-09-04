using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet.Query;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Text
{
    class ParagraphFormatCellsReader : ReaderMultiRow<Text.ParagraphFormatCells>
    {
        public SubQueryColumn Bullet { get; set; }
        public SubQueryColumn BulletFont { get; set; }
        public SubQueryColumn BulletFontSize { get; set; }
        public SubQueryColumn BulletString { get; set; }
        public SubQueryColumn Flags { get; set; }
        public SubQueryColumn HorizontalAlign { get; set; }
        public SubQueryColumn IndentFirst { get; set; }
        public SubQueryColumn IndentLeft { get; set; }
        public SubQueryColumn IndentRight { get; set; }
        public SubQueryColumn LocalizeBulletFont { get; set; }
        public SubQueryColumn SpaceAfter { get; set; }
        public SubQueryColumn SpaceBefore { get; set; }
        public SubQueryColumn SpaceLine { get; set; }
        public SubQueryColumn TextPosAfterBullet { get; set; }

        public ParagraphFormatCellsReader()
        {
            var sec = this.query.AddSubQuery(IVisio.VisSectionIndices.visSectionParagraph);
            this.Bullet = sec.AddCell(SrcConstants.ParaBullet, nameof(SrcConstants.ParaBullet));
            this.BulletFont = sec.AddCell(SrcConstants.ParaBulletFont, nameof(SrcConstants.ParaBulletFont));
            this.BulletFontSize = sec.AddCell(SrcConstants.ParaBulletFontSize, nameof(SrcConstants.ParaBulletFontSize));
            this.BulletString = sec.AddCell(SrcConstants.ParaBulletString, nameof(SrcConstants.ParaBulletString));
            this.Flags = sec.AddCell(SrcConstants.ParaFlags, nameof(SrcConstants.ParaFlags));
            this.HorizontalAlign = sec.AddCell(SrcConstants.ParaHorizontalAlign, nameof(SrcConstants.ParaHorizontalAlign));
            this.IndentFirst = sec.AddCell(SrcConstants.ParaIndentFirst, nameof(SrcConstants.ParaIndentFirst));
            this.IndentLeft = sec.AddCell(SrcConstants.ParaIndentLeft, nameof(SrcConstants.ParaIndentLeft));
            this.IndentRight = sec.AddCell(SrcConstants.ParaIndentRight, nameof(SrcConstants.ParaIndentRight));
            this.LocalizeBulletFont = sec.AddCell(SrcConstants.ParaLocalizeBulletFont, nameof(SrcConstants.ParaLocalizeBulletFont));
            this.SpaceAfter = sec.AddCell(SrcConstants.ParaSpacingAfter, nameof(SrcConstants.ParaSpacingAfter));
            this.SpaceBefore = sec.AddCell(SrcConstants.ParaSpacingBefore, nameof(SrcConstants.ParaSpacingBefore));
            this.SpaceLine = sec.AddCell(SrcConstants.ParaSpacingLine, nameof(SrcConstants.ParaSpacingLine));
            this.TextPosAfterBullet = sec.AddCell(SrcConstants.ParaTextPosAfterBullet, nameof(SrcConstants.ParaTextPosAfterBullet));
        }

        public override Text.ParagraphFormatCells CellDataToCellGroup(VisioAutomation.Utilities.ArraySegment<ShapeSheet.CellData> row)
        {
            var cells = new Text.ParagraphFormatCells();
            cells.IndentFirst = row[this.IndentFirst];
            cells.IndentLeft = row[this.IndentLeft];
            cells.IndentRight = row[this.IndentRight];
            cells.SpacingAfter = row[this.SpaceAfter];
            cells.SpacingBefore = row[this.SpaceBefore];
            cells.SpacingLine = row[this.SpaceLine];
            cells.HorizontalAlign = row[this.HorizontalAlign];
            cells.Bullet = row[this.Bullet];
            cells.BulletFont = row[this.BulletFont];
            cells.BulletFontSize = row[this.BulletFontSize];
            cells.LocalizeBulletFont = row[this.LocalizeBulletFont];
            cells.TextPosAfterBullet = row[this.TextPosAfterBullet];
            cells.Flags = row[this.Flags];
            cells.BulletString = row[this.BulletString];

            return cells;
        }
    }
}