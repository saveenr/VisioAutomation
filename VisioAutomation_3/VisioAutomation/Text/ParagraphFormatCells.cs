using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using VisioAutomation.Extensions;

namespace VisioAutomation.Text
{
    public class ParagraphFormatCells : VA.ShapeSheet.CellSectionDataGroup
    {
        ////public string BulletString;
        public VA.ShapeSheet.CellData<double> IndentFirst { get; set; }
        public VA.ShapeSheet.CellData<double> IndentRight { get; set; }
        public VA.ShapeSheet.CellData<double> IndentLeft { get; set; }
        public VA.ShapeSheet.CellData<double> SpacingBefore { get; set; }
        public VA.ShapeSheet.CellData<double> SpacingAfter { get; set; }
        public VA.ShapeSheet.CellData<double> SpacingLine { get; set; }
        public VA.ShapeSheet.CellData<int> HorizontalAlign { get; set; }
        public VA.ShapeSheet.CellData<int> BulletIndex { get; set; }
        public VA.ShapeSheet.CellData<int> BulletFont { get; set; }
        public VA.ShapeSheet.CellData<int> BulletFontSize { get; set; }

        protected override void _Apply(VA.ShapeSheet.CellSectionDataGroup.ApplyFormula func, short row)
        {
            func(VA.ShapeSheet.SRCConstants.Para_IndLeft.ForRow(row), this.IndentLeft.Formula);
            func(VA.ShapeSheet.SRCConstants.Para_IndFirst.ForRow(row), this.IndentFirst.Formula);
            func(VA.ShapeSheet.SRCConstants.Para_IndRight.ForRow(row), this.IndentRight.Formula);
            func(VA.ShapeSheet.SRCConstants.Para_SpAfter.ForRow(row), this.SpacingAfter.Formula);
            func(VA.ShapeSheet.SRCConstants.Para_SpBefore.ForRow(row), this.SpacingBefore.Formula);
            func(VA.ShapeSheet.SRCConstants.Para_SpLine.ForRow(row), this.SpacingLine.Formula);
            func(VA.ShapeSheet.SRCConstants.Para_HorzAlign.ForRow(row), this.HorizontalAlign.Formula);
            func(VA.ShapeSheet.SRCConstants.Para_BulletFont.ForRow(row), this.BulletFont.Formula);
            func(VA.ShapeSheet.SRCConstants.Para_Bullet.ForRow(row), this.BulletIndex.Formula);
            func(VA.ShapeSheet.SRCConstants.Para_BulletFontSize.ForRow(row), this.BulletFontSize.Formula);
        }

        internal static IList<List<ParagraphFormatCells>> GetCells(IVisio.Page page, IList<int> shapeids)
        {
            var query = new ParagraphFormatQuery();
            return VA.ShapeSheet.CellSectionDataGroup._GetCells(page, shapeids, query, get_cells_from_row);
        }

        internal static IList<ParagraphFormatCells> GetCells(IVisio.Shape shape)
        {
            var query = new ParagraphFormatQuery();
            return VA.ShapeSheet.CellSectionDataGroup._GetCells(shape, query, get_cells_from_row);
        }

        private static ParagraphFormatCells get_cells_from_row(ParagraphFormatQuery query, VA.ShapeSheet.Query.QueryDataSet<double> qds, int row)
        {
            var cells = new ParagraphFormatCells();
            cells.IndentFirst = qds.GetItem(row, query.IndentFirst);
            cells.IndentLeft = qds.GetItem(row, query.IndentLeft);
            cells.IndentRight = qds.GetItem(row, query.IndentRight);
            cells.SpacingAfter = qds.GetItem(row, query.SpaceAfter);
            cells.SpacingBefore = qds.GetItem(row, query.SpaceBefore);
            cells.SpacingLine = qds.GetItem(row, query.SpaceLine);
            cells.HorizontalAlign = qds.GetItem(row, query.HorzAlign).ToInt();
            cells.BulletIndex = qds.GetItem(row, query.BulletIndex).ToInt();
            cells.BulletFont = qds.GetItem(row, query.BulletFont).ToInt();
            cells.BulletFontSize = qds.GetItem(row, query.BulletFontSize).ToInt();

            return cells;
        }

        class ParagraphFormatQuery : VA.ShapeSheet.Query.SectionQuery
        {
            public VA.ShapeSheet.Query.SectionQueryColumn BulletIndex { get; set; }
            public VA.ShapeSheet.Query.SectionQueryColumn BulletFont { get; set; }
            public VA.ShapeSheet.Query.SectionQueryColumn BulletFontSize { get; set; }
            public VA.ShapeSheet.Query.SectionQueryColumn BulletString { get; set; }
            public VA.ShapeSheet.Query.SectionQueryColumn Flags { get; set; }
            public VA.ShapeSheet.Query.SectionQueryColumn HorzAlign { get; set; }
            public VA.ShapeSheet.Query.SectionQueryColumn IndentFirst { get; set; }
            public VA.ShapeSheet.Query.SectionQueryColumn IndentLeft { get; set; }
            public VA.ShapeSheet.Query.SectionQueryColumn IndentRight { get; set; }
            public VA.ShapeSheet.Query.SectionQueryColumn LocalizeBulletFont { get; set; }
            public VA.ShapeSheet.Query.SectionQueryColumn SpaceAfter { get; set; }
            public VA.ShapeSheet.Query.SectionQueryColumn SpaceBefore { get; set; }
            public VA.ShapeSheet.Query.SectionQueryColumn SpaceLine { get; set; }
            public VA.ShapeSheet.Query.SectionQueryColumn TextPosAfterBullet { get; set; }

            public ParagraphFormatQuery() :
                base(IVisio.VisSectionIndices.visSectionParagraph)
            {
                BulletIndex = this.AddColumn(VA.ShapeSheet.SRCConstants.Para_Bullet, "BulletIndex");
                BulletFont = this.AddColumn(VA.ShapeSheet.SRCConstants.Para_BulletFont, "BulletFont");
                BulletFontSize = this.AddColumn(VA.ShapeSheet.SRCConstants.Para_BulletFontSize, "BulletFontSize");
                BulletString = this.AddColumn(VA.ShapeSheet.SRCConstants.Para_BulletStr, "BulletString");
                Flags = this.AddColumn(VA.ShapeSheet.SRCConstants.Para_Flags, "Flags");
                HorzAlign = this.AddColumn(VA.ShapeSheet.SRCConstants.Para_HorzAlign, "HorzAlign");
                IndentFirst = this.AddColumn(VA.ShapeSheet.SRCConstants.Para_IndFirst, "IndentFirst");
                IndentLeft = this.AddColumn(VA.ShapeSheet.SRCConstants.Para_IndLeft, "IndentLeft");
                IndentRight = this.AddColumn(VA.ShapeSheet.SRCConstants.Para_IndRight, "IndentRight");
                LocalizeBulletFont = this.AddColumn(VA.ShapeSheet.SRCConstants.Para_LocalizeBulletFont, "LocalizeBulletFont");
                SpaceAfter = this.AddColumn(VA.ShapeSheet.SRCConstants.Para_SpAfter, "SpaceAfter");
                SpaceBefore = this.AddColumn(VA.ShapeSheet.SRCConstants.Para_SpBefore, "SpaceBefore");
                SpaceLine = this.AddColumn(VA.ShapeSheet.SRCConstants.Para_SpLine, "SpaceLine");
                TextPosAfterBullet = this.AddColumn(VA.ShapeSheet.SRCConstants.Para_TextPosAfterBullet, "TextPosAfterBullet");
            }
        }
    }
} 