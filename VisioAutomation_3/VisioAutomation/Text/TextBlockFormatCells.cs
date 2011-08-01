using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.Text
{
    public class TextBlockFormatCells : VA.ShapeSheet.CellDataGroup
    {
        public VA.ShapeSheet.CellData<double> BottomMargin { get; set; }
        public VA.ShapeSheet.CellData<double> LeftMargin { get; set; }
        public VA.ShapeSheet.CellData<double> RightMargin { get; set; }
        public VA.ShapeSheet.CellData<double> TopMargin { get; set; }

        public VA.ShapeSheet.CellData<double> DefaultTabStop { get; set; }
        
        public VA.ShapeSheet.CellData<int> TextBkgnd { get; set; }
        public VA.ShapeSheet.CellData<double> TextBkgndTrans { get; set; }
        
        public VA.ShapeSheet.CellData<int> TextDirection { get; set; }
        
        public VA.ShapeSheet.CellData<int> VerticalAlign { get; set; }

        protected override void _Apply(VA.ShapeSheet.CellDataGroup.ApplyFormula func)
        {
            func(VA.ShapeSheet.SRCConstants.BottomMargin, this.BottomMargin.Formula);
            func(VA.ShapeSheet.SRCConstants.LeftMargin, this.LeftMargin.Formula);
            func(VA.ShapeSheet.SRCConstants.RightMargin, this.RightMargin.Formula);
            func(VA.ShapeSheet.SRCConstants.TopMargin, this.TopMargin.Formula);
            func(VA.ShapeSheet.SRCConstants.DefaultTabStop, this.DefaultTabStop.Formula);
            func(VA.ShapeSheet.SRCConstants.TextBkgnd, this.TextBkgnd.Formula);
            func(VA.ShapeSheet.SRCConstants.TextBkgndTrans, this.TextBkgndTrans.Formula);
            func(VA.ShapeSheet.SRCConstants.TextDirection, this.TextDirection.Formula);
            func(VA.ShapeSheet.SRCConstants.VerticalAlign, this.VerticalAlign.Formula);
        }

        internal static IList<TextBlockFormatCells> GetCells(IVisio.Page page, IList<int> shapeids)
        {
            var query = new TextBlockFormatQuery();
            return VA.ShapeSheet.CellDataGroup._GetCells(page, shapeids, query, get_cells_from_row);
        }

        internal static TextBlockFormatCells GetCells(IVisio.Shape shape)
        {
            var query = new TextBlockFormatQuery();
            return VA.ShapeSheet.CellDataGroup._GetCells(shape, query, get_cells_from_row);
        }

        private static TextBlockFormatCells get_cells_from_row(TextBlockFormatQuery query, VA.ShapeSheet.Query.QueryDataSet<double> qds, int row)
        {
            var cells = new TextBlockFormatCells();
            cells.BottomMargin = qds.GetItem(row, query.BottomMargin);
            cells.LeftMargin= qds.GetItem(row, query.LeftMargin);
            cells.RightMargin = qds.GetItem(row, query.RightMargin);
            cells.TopMargin = qds.GetItem(row, query.TopMargin);
            cells.DefaultTabStop = qds.GetItem(row, query.DefaultTabStop);
            cells.TextBkgnd = qds.GetItem(row, query.TextBkgnd, v => (int)v);
            cells.TextBkgndTrans = qds.GetItem(row, query.TextBkgndTrans);
            cells.TextDirection = qds.GetItem(row, query.TextDirection, v => (int)v);
            cells.VerticalAlign = qds.GetItem(row, query.VerticalAlign, v => (int)v);
            return cells;
        }
    }
}