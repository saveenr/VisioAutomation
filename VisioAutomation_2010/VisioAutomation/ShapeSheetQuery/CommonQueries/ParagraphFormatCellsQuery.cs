using SRCCON=VisioAutomation.ShapeSheet.SRCConstants;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheetQuery.CommonQueries
{
    class ParagraphFormatCellsQuery : Query
    {
        public ColumnCellIndex Bullet { get; set; }
        public ColumnCellIndex BulletFont { get; set; }
        public ColumnCellIndex BulletFontSize { get; set; }
        public ColumnCellIndex BulletString { get; set; } // NOTE: This is never used
        public ColumnCellIndex Flags { get; set; }
        public ColumnCellIndex HorzAlign { get; set; }
        public ColumnCellIndex IndentFirst { get; set; }
        public ColumnCellIndex IndentLeft { get; set; }
        public ColumnCellIndex IndentRight { get; set; }
        public ColumnCellIndex LocalizeBulletFont { get; set; }
        public ColumnCellIndex SpaceAfter { get; set; }
        public ColumnCellIndex SpaceBefore { get; set; }
        public ColumnCellIndex SpaceLine { get; set; }
        public ColumnCellIndex TextPosAfterBullet { get; set; }

        public ParagraphFormatCellsQuery()
        {
            var sec = this.AddSection(IVisio.VisSectionIndices.visSectionParagraph);
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

        public Text.ParagraphCells GetCells(ShapeSheet.CellData<double>[] row)
        {
            var cells = new Text.ParagraphCells();
            cells.IndentFirst = row[this.IndentFirst];
            cells.IndentLeft = row[this.IndentLeft];
            cells.IndentRight = row[this.IndentRight];
            cells.SpacingAfter = row[this.SpaceAfter];
            cells.SpacingBefore = row[this.SpaceBefore];
            cells.SpacingLine = row[this.SpaceLine];
            cells.HorizontalAlign = Extensions.CellDataMethods.ToInt(row[this.HorzAlign]);
            cells.Bullet = Extensions.CellDataMethods.ToInt(row[this.Bullet]);
            cells.BulletFont = Extensions.CellDataMethods.ToInt(row[this.BulletFont]);
            cells.BulletFontSize = Extensions.CellDataMethods.ToInt(row[this.BulletFontSize]);
            cells.LocBulletFont = Extensions.CellDataMethods.ToInt(row[this.LocalizeBulletFont]);
            cells.TextPosAfterBullet = row[this.TextPosAfterBullet];
            cells.Flags = Extensions.CellDataMethods.ToInt(row[this.Flags]);
            cells.BulletString = ""; // TODO: Figure out some way of getting this

            return cells;
        }
    }
}