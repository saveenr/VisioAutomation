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
            var sec = this.query.AddSubQuery(IVisio.VisSectionIndices.visSectionParagraph);
            this.Bullet = sec.AddColumn(SrcConstants.ParaBullet, nameof(SrcConstants.ParaBullet));
            this.BulletFont = sec.AddColumn(SrcConstants.ParaBulletFont, nameof(SrcConstants.ParaBulletFont));
            this.BulletFontSize = sec.AddColumn(SrcConstants.ParaBulletFontSize, nameof(SrcConstants.ParaBulletFontSize));
            this.BulletString = sec.AddColumn(SrcConstants.ParaBulletString, nameof(SrcConstants.ParaBulletString));
            this.Flags = sec.AddColumn(SrcConstants.ParaFlags, nameof(SrcConstants.ParaFlags));
            this.HorizontalAlign = sec.AddColumn(SrcConstants.ParaHorizontalAlign, nameof(SrcConstants.ParaHorizontalAlign));
            this.IndentFirst = sec.AddColumn(SrcConstants.ParaIndentFirst, nameof(SrcConstants.ParaIndentFirst));
            this.IndentLeft = sec.AddColumn(SrcConstants.ParaIndentLeft, nameof(SrcConstants.ParaIndentLeft));
            this.IndentRight = sec.AddColumn(SrcConstants.ParaIndentRight, nameof(SrcConstants.ParaIndentRight));
            this.LocalizeBulletFont = sec.AddColumn(SrcConstants.ParaLocalizeBulletFont, nameof(SrcConstants.ParaLocalizeBulletFont));
            this.SpaceAfter = sec.AddColumn(SrcConstants.ParaSpacingAfter, nameof(SrcConstants.ParaSpacingAfter));
            this.SpaceBefore = sec.AddColumn(SrcConstants.ParaSpacingBefore, nameof(SrcConstants.ParaSpacingBefore));
            this.SpaceLine = sec.AddColumn(SrcConstants.ParaSpacingLine, nameof(SrcConstants.ParaSpacingLine));
            this.TextPosAfterBullet = sec.AddColumn(SrcConstants.ParaTextPosAfterBullet, nameof(SrcConstants.ParaTextPosAfterBullet));
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