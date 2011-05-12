using VA=VisioAutomation;
using System;
using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Format
{
    public partial class StyleCells : VA.ShapeSheet.CellDataGroup
    {
        public VA.ShapeSheet.CellData<int> EnableFillProps { get; set; }
        public VA.ShapeSheet.CellData<int> EnableLineProps { get; set; }
        public VA.ShapeSheet.CellData<int> EnableTextProps { get; set; }
        public VA.ShapeSheet.CellData<bool> HideForApply { get; set; }

        protected override void _Apply(VA.ShapeSheet.CellDataGroup.ApplyFormula func)
        {
            func(ShapeSheet.SRCConstants.EnableFillProps, this.EnableFillProps.Formula);
            func(ShapeSheet.SRCConstants.EnableLineProps, this.EnableLineProps.Formula);
            func(ShapeSheet.SRCConstants.EnableTextProps, this.EnableTextProps.Formula);
            func(ShapeSheet.SRCConstants.HideForApply, this.HideForApply.Formula);
        }

        private static StyleCells get_cells_from_row(StyleQuery query, VA.ShapeSheet.Query.QueryDataSet<double> qds, int row)
        {
            var cells = new StyleCells();
            cells.EnableFillProps = qds.GetItem(row, query.EnableFillProps, v => (int)v);
            cells.EnableLineProps = qds.GetItem(row, query.EnableLineProps, v => (int)v);
            cells.EnableTextProps = qds.GetItem(row, query.EnableTextProps, v => (int)v);
            cells.HideForApply = qds.GetItem(row, query.HideForApply, v => VA.Convert.DoubleToBool(v));
            return cells;
        }

        internal static IList<StyleCells> GetCells(IVisio.Page page, IList<int> shapeids)
        {
            var query = new StyleQuery();
            return VA.ShapeSheet.CellDataGroup._GetCells(page, shapeids, query, get_cells_from_row);
        }

        internal static StyleCells GetCells(IVisio.Shape shape)
        {
            var query = new StyleQuery();
            return VA.ShapeSheet.CellDataGroup._GetCells(shape, query, get_cells_from_row);
        }

    }
}