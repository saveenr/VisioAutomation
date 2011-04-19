using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

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
        public VA.ShapeSheet.CellData<int> BulletSize { get; set; }

        protected override void _Apply(VA.ShapeSheet.CellSectionDataGroup.ApplyFormula func, short row)
        {
            func(VA.ShapeSheet.SRCConstants.Para_IndLeft.ForRow(row), this.IndentLeft.Formula);
            func(VA.ShapeSheet.SRCConstants.Para_IndFirst.ForRow(row), this.IndentFirst.Formula);
            func(VA.ShapeSheet.SRCConstants.Para_IndRight.ForRow(row), this.IndentRight.Formula);
            func(VA.ShapeSheet.SRCConstants.Para_SpAfter.ForRow(row), this.SpacingAfter.Formula);
            func(VA.ShapeSheet.SRCConstants.Para_SpBefore.ForRow(row), this.SpacingBefore.Formula);
            func(VA.ShapeSheet.SRCConstants.Para_SpLine.ForRow(row), this.SpacingLine.Formula);
            func(VA.ShapeSheet.SRCConstants.Para_HAlign.ForRow(row), this.HorizontalAlign.Formula);
            func(VA.ShapeSheet.SRCConstants.Para_BulletFont.ForRow(row), this.BulletFont.Formula);
            func(VA.ShapeSheet.SRCConstants.Para_BulletIndex.ForRow(row), this.BulletIndex.Formula);
            func(VA.ShapeSheet.SRCConstants.Para_BulletSize.ForRow(row), this.BulletSize.Formula);
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
            cells.HorizontalAlign = qds.GetItem(row, query.HorzAlign, v => (int)v);
            cells.BulletIndex = qds.GetItem(row, query.BulletIndex, v => (int)v);
            cells.BulletFont = qds.GetItem(row, query.BulletFont, v => (int)v);
            cells.BulletSize = qds.GetItem(row, query.BulletFontSize, v => (int)v);

            return cells;
        }
    }
}