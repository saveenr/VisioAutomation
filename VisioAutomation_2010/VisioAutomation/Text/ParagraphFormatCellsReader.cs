using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet.Query;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Text
{
    class ParagraphFormatCellsReader : MultiRowReader<Text.ParagraphCells>
    {
        public SubQueryColumn Bullet { get; set; }
        public SubQueryColumn BulletFont { get; set; }
        public SubQueryColumn BulletFontSize { get; set; }
        public SubQueryColumn BulletString { get; set; }
        public SubQueryColumn Flags { get; set; }
        public SubQueryColumn HorzAlign { get; set; }
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
            this.Bullet = sec.AddCell(SrcConstants.Para_Bullet, nameof(SrcConstants.Para_Bullet));
            this.BulletFont = sec.AddCell(SrcConstants.Para_BulletFont, nameof(SrcConstants.Para_BulletFont));
            this.BulletFontSize = sec.AddCell(SrcConstants.Para_BulletFontSize, nameof(SrcConstants.Para_BulletFontSize));
            this.BulletString = sec.AddCell(SrcConstants.Para_BulletStr, nameof(SrcConstants.Para_BulletStr));
            this.Flags = sec.AddCell(SrcConstants.Para_Flags, nameof(SrcConstants.Para_Flags));
            this.HorzAlign = sec.AddCell(SrcConstants.Para_HorzAlign, nameof(SrcConstants.Para_HorzAlign));
            this.IndentFirst = sec.AddCell(SrcConstants.Para_IndFirst, nameof(SrcConstants.Para_IndFirst));
            this.IndentLeft = sec.AddCell(SrcConstants.Para_IndLeft, nameof(SrcConstants.Para_IndLeft));
            this.IndentRight = sec.AddCell(SrcConstants.Para_IndRight, nameof(SrcConstants.Para_IndRight));
            this.LocalizeBulletFont = sec.AddCell(SrcConstants.Para_LocalizeBulletFont, nameof(SrcConstants.Para_LocalizeBulletFont));
            this.SpaceAfter = sec.AddCell(SrcConstants.Para_SpAfter, nameof(SrcConstants.Para_SpAfter));
            this.SpaceBefore = sec.AddCell(SrcConstants.Para_SpBefore, nameof(SrcConstants.Para_SpBefore));
            this.SpaceLine = sec.AddCell(SrcConstants.Para_SpLine, nameof(SrcConstants.Para_SpLine));
            this.TextPosAfterBullet = sec.AddCell(SrcConstants.Para_TextPosAfterBullet, nameof(SrcConstants.Para_TextPosAfterBullet));
        }

        public override Text.ParagraphCells CellDataToCellGroup(VisioAutomation.Utilities.ArraySegment<ShapeSheet.CellData> row)
        {
            var cells = new Text.ParagraphCells();
            cells.IndentFirst = row[this.IndentFirst];
            cells.IndentLeft = row[this.IndentLeft];
            cells.IndentRight = row[this.IndentRight];
            cells.SpacingAfter = row[this.SpaceAfter];
            cells.SpacingBefore = row[this.SpaceBefore];
            cells.SpacingLine = row[this.SpaceLine];
            cells.HorizontalAlign = row[this.HorzAlign];
            cells.Bullet = row[this.Bullet];
            cells.BulletFont = row[this.BulletFont];
            cells.BulletFontSize = row[this.BulletFontSize];
            cells.LocBulletFont = row[this.LocalizeBulletFont];
            cells.TextPosAfterBullet = row[this.TextPosAfterBullet];
            cells.Flags = row[this.Flags];
            cells.BulletString = row[this.BulletString];

            return cells;
        }
    }
}