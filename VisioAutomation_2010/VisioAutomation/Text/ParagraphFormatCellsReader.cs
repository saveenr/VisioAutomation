using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet.Query;
using SRCCON=VisioAutomation.ShapeSheet.SRCConstants;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Text
{
    class ParagraphFormatCellsReader : MultiRowReader<Text.ParagraphCells>
    {
        public ColumnSubQuery Bullet { get; set; }
        public ColumnSubQuery BulletFont { get; set; }
        public ColumnSubQuery BulletFontSize { get; set; }
        public ColumnSubQuery BulletString { get; set; }
        public ColumnSubQuery Flags { get; set; }
        public ColumnSubQuery HorzAlign { get; set; }
        public ColumnSubQuery IndentFirst { get; set; }
        public ColumnSubQuery IndentLeft { get; set; }
        public ColumnSubQuery IndentRight { get; set; }
        public ColumnSubQuery LocalizeBulletFont { get; set; }
        public ColumnSubQuery SpaceAfter { get; set; }
        public ColumnSubQuery SpaceBefore { get; set; }
        public ColumnSubQuery SpaceLine { get; set; }
        public ColumnSubQuery TextPosAfterBullet { get; set; }

        public ParagraphFormatCellsReader()
        {
            var sec = this.query.AddSubQuery(IVisio.VisSectionIndices.visSectionParagraph);
            this.Bullet = sec.AddCell(SRCCON.Para_Bullet, nameof(SRCCON.Para_Bullet));
            this.BulletFont = sec.AddCell(SRCCON.Para_BulletFont, nameof(SRCCON.Para_BulletFont));
            this.BulletFontSize = sec.AddCell(SRCCON.Para_BulletFontSize, nameof(SRCCON.Para_BulletFontSize));
            this.BulletString = sec.AddCell(SRCCON.Para_BulletStr, nameof(SRCCON.Para_BulletStr));
            this.Flags = sec.AddCell(SRCCON.Para_Flags, nameof(SRCCON.Para_Flags));
            this.HorzAlign = sec.AddCell(SRCCON.Para_HorzAlign, nameof(SRCCON.Para_HorzAlign));
            this.IndentFirst = sec.AddCell(SRCCON.Para_IndFirst, nameof(SRCCON.Para_IndFirst));
            this.IndentLeft = sec.AddCell(SRCCON.Para_IndLeft, nameof(SRCCON.Para_IndLeft));
            this.IndentRight = sec.AddCell(SRCCON.Para_IndRight, nameof(SRCCON.Para_IndRight));
            this.LocalizeBulletFont = sec.AddCell(SRCCON.Para_LocalizeBulletFont, nameof(SRCCON.Para_LocalizeBulletFont));
            this.SpaceAfter = sec.AddCell(SRCCON.Para_SpAfter, nameof(SRCCON.Para_SpAfter));
            this.SpaceBefore = sec.AddCell(SRCCON.Para_SpBefore, nameof(SRCCON.Para_SpBefore));
            this.SpaceLine = sec.AddCell(SRCCON.Para_SpLine, nameof(SRCCON.Para_SpLine));
            this.TextPosAfterBullet = sec.AddCell(SRCCON.Para_TextPosAfterBullet, nameof(SRCCON.Para_TextPosAfterBullet));
        }

        public override Text.ParagraphCells CellDataToCellGroup(ShapeSheet.CellData[] row)
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