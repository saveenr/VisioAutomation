namespace VisioAutomation.ShapeSheet.Query.Common
{
    class ParagraphFormatCellsQuery : CellQuery
    {
        public Query.CellColumn Bullet { get; set; }
        public Query.CellColumn BulletFont { get; set; }
        public Query.CellColumn BulletFontSize { get; set; }
        public Query.CellColumn BulletString { get; set; } // NOTE: This is never used
        public Query.CellColumn Flags { get; set; }
        public Query.CellColumn HorzAlign { get; set; }
        public Query.CellColumn IndentFirst { get; set; }
        public Query.CellColumn IndentLeft { get; set; }
        public Query.CellColumn IndentRight { get; set; }
        public Query.CellColumn LocalizeBulletFont { get; set; }
        public Query.CellColumn SpaceAfter { get; set; }
        public Query.CellColumn SpaceBefore { get; set; }
        public Query.CellColumn SpaceLine { get; set; }
        public Query.CellColumn TextPosAfterBullet { get; set; }

        public ParagraphFormatCellsQuery()
        {
            var sec = this.AddSection(Microsoft.Office.Interop.Visio.VisSectionIndices.visSectionParagraph);
            this.Bullet = sec.AddCell(ShapeSheet.SRCConstants.Para_Bullet, nameof(ShapeSheet.SRCConstants.Para_Bullet));
            this.BulletFont = sec.AddCell(ShapeSheet.SRCConstants.Para_BulletFont, nameof(ShapeSheet.SRCConstants.Para_BulletFont));
            this.BulletFontSize = sec.AddCell(ShapeSheet.SRCConstants.Para_BulletFontSize, nameof(ShapeSheet.SRCConstants.Para_BulletFontSize));
            this.BulletString = sec.AddCell(ShapeSheet.SRCConstants.Para_BulletStr, nameof(ShapeSheet.SRCConstants.Para_BulletStr));
            this.Flags = sec.AddCell(ShapeSheet.SRCConstants.Para_Flags, nameof(ShapeSheet.SRCConstants.Para_Flags));
            this.HorzAlign = sec.AddCell(ShapeSheet.SRCConstants.Para_HorzAlign, nameof(ShapeSheet.SRCConstants.Para_HorzAlign));
            this.IndentFirst = sec.AddCell(ShapeSheet.SRCConstants.Para_IndFirst, nameof(ShapeSheet.SRCConstants.Para_IndFirst));
            this.IndentLeft = sec.AddCell(ShapeSheet.SRCConstants.Para_IndLeft, nameof(ShapeSheet.SRCConstants.Para_IndLeft));
            this.IndentRight = sec.AddCell(ShapeSheet.SRCConstants.Para_IndRight, nameof(ShapeSheet.SRCConstants.Para_IndRight));
            this.LocalizeBulletFont = sec.AddCell(ShapeSheet.SRCConstants.Para_LocalizeBulletFont, nameof(ShapeSheet.SRCConstants.Para_LocalizeBulletFont));
            this.SpaceAfter = sec.AddCell(ShapeSheet.SRCConstants.Para_SpAfter, nameof(ShapeSheet.SRCConstants.Para_SpAfter));
            this.SpaceBefore = sec.AddCell(ShapeSheet.SRCConstants.Para_SpBefore, nameof(ShapeSheet.SRCConstants.Para_SpBefore));
            this.SpaceLine = sec.AddCell(ShapeSheet.SRCConstants.Para_SpLine, nameof(ShapeSheet.SRCConstants.Para_SpLine));
            this.TextPosAfterBullet = sec.AddCell(ShapeSheet.SRCConstants.Para_TextPosAfterBullet, nameof(ShapeSheet.SRCConstants.Para_TextPosAfterBullet));


        }

        public VisioAutomation.Text.ParagraphCells GetCells(System.Collections.Generic.IList<ShapeSheet.CellData<double>> row)
        {
            var cells = new VisioAutomation.Text.ParagraphCells();
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