using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet.Query;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Text
{
    class ParagraphFormatCellsReader : ReaderMultiRow<Text.ParagraphFormatCells>
    {
        public SectionQueryColumn Bullet { get; set; }
        public SectionQueryColumn BulletFont { get; set; }
        public SectionQueryColumn BulletFontSize { get; set; }
        public SectionQueryColumn BulletString { get; set; }
        public SectionQueryColumn Flags { get; set; }
        public SectionQueryColumn HorizontalAlign { get; set; }
        public SectionQueryColumn IndentFirst { get; set; }
        public SectionQueryColumn IndentLeft { get; set; }
        public SectionQueryColumn IndentRight { get; set; }
        public SectionQueryColumn LocalizeBulletFont { get; set; }
        public SectionQueryColumn SpaceAfter { get; set; }
        public SectionQueryColumn SpaceBefore { get; set; }
        public SectionQueryColumn SpaceLine { get; set; }
        public SectionQueryColumn TextPosAfterBullet { get; set; }

        public ParagraphFormatCellsReader()
        {
            var sec = this.query.SectionQueries.Add(IVisio.VisSectionIndices.visSectionParagraph);
            this.Bullet = sec.Columns.Add(SrcConstants.ParaBullet, nameof(this.Bullet));
            this.BulletFont = sec.Columns.Add(SrcConstants.ParaBulletFont, nameof(this.BulletFont));
            this.BulletFontSize = sec.Columns.Add(SrcConstants.ParaBulletFontSize, nameof(this.BulletFontSize));
            this.BulletString = sec.Columns.Add(SrcConstants.ParaBulletString, nameof(this.BulletString));
            this.Flags = sec.Columns.Add(SrcConstants.ParaFlags, nameof(this.Flags));
            this.HorizontalAlign = sec.Columns.Add(SrcConstants.ParaHorizontalAlign, nameof(this.HorizontalAlign));
            this.IndentFirst = sec.Columns.Add(SrcConstants.ParaIndentFirst, nameof(this.IndentFirst));
            this.IndentLeft = sec.Columns.Add(SrcConstants.ParaIndentLeft, nameof(this.IndentLeft));
            this.IndentRight = sec.Columns.Add(SrcConstants.ParaIndentRight, nameof(this.IndentRight));
            this.LocalizeBulletFont = sec.Columns.Add(SrcConstants.ParaLocalizeBulletFont, nameof(this.LocalizeBulletFont));
            this.SpaceAfter = sec.Columns.Add(SrcConstants.ParaSpacingAfter, nameof(this.SpaceAfter));
            this.SpaceBefore = sec.Columns.Add(SrcConstants.ParaSpacingBefore, nameof(this.SpaceBefore));
            this.SpaceLine = sec.Columns.Add(SrcConstants.ParaSpacingLine, nameof(this.SpaceLine));
            this.TextPosAfterBullet = sec.Columns.Add(SrcConstants.ParaTextPosAfterBullet, nameof(this.TextPosAfterBullet));
        }

        public override Text.ParagraphFormatCells CellDataToCellGroup(VisioAutomation.Utilities.ArraySegment<string> row)
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