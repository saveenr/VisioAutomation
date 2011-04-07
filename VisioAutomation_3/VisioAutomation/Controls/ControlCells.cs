using System.Collections.Generic;
using VA=VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Controls
{
    public class ControlCells
    {
        public VA.ShapeSheet.CellData<int> CanGlue { get; set; }
        public VA.ShapeSheet.CellData<int> Tip { get; set; }
        public VA.ShapeSheet.CellData<double> X { get; set; }
        public VA.ShapeSheet.CellData<double> Y { get; set; }
        public VA.ShapeSheet.CellData<int> YBehavior { get; set; }
        public VA.ShapeSheet.CellData<int> XBehavior { get; set; }
        public VA.ShapeSheet.CellData<int> XDynamics { get; set; }
        public VA.ShapeSheet.CellData<int> YDynamics { get; set; }

        public void Apply(VA.ShapeSheet.Update.SRCUpdate update, short row)
        {
            this._Apply((src, f) => update.SetFormulaIgnoreNull(src, f), row);
        }

        public void Apply(VA.ShapeSheet.Update.SIDSRCUpdate update, short id, short row)
        {
            this._Apply((src, f) => update.SetFormulaIgnoreNull(id, src, f), row);
        }

        internal void _Apply(System.Action<VA.ShapeSheet.SRC, VA.ShapeSheet.FormulaLiteral> func, short row)
        {
            func(VA.ShapeSheet.SRCConstants.Controls_CanGlue.ForRow(row), this.CanGlue.Formula);
            func(VA.ShapeSheet.SRCConstants.Controls_Tip.ForRow(row), this.Tip.Formula);
            func(VA.ShapeSheet.SRCConstants.Controls_X.ForRow(row), this.X.Formula);
            func(VA.ShapeSheet.SRCConstants.Controls_Y.ForRow(row), this.Y.Formula);
            func(VA.ShapeSheet.SRCConstants.Controls_YCon.ForRow(row), this.YBehavior.Formula);
            func(VA.ShapeSheet.SRCConstants.Controls_XCon.ForRow(row), this.XBehavior.Formula);
            func(VA.ShapeSheet.SRCConstants.Controls_XDyn.ForRow(row), this.XDynamics.Formula);
            func(VA.ShapeSheet.SRCConstants.Controls_YDyn.ForRow(row), this.YDynamics.Formula);
        }

        private static ControlCells get_cells_from_row(ControlQuery query, VA.ShapeSheet.Query.QueryDataSet<double> qds, int row)
        {
            var cells = new ControlCells();
            cells.CanGlue = qds.GetItem(row, query.CanGlue, v => (int)v);
            cells.Tip = qds.GetItem(row, query.Tip, v => (int)v);
            cells.X = qds.GetItem(row, query.X);
            cells.Y = qds.GetItem(row, query.Y);
            cells.YBehavior = qds.GetItem(row, query.YBehavior, v => (int)v);
            cells.XBehavior = qds.GetItem(row, query.XBehavior, v => (int)v);
            cells.XDynamics = qds.GetItem(row, query.XDynamics, v => (int)v);
            cells.YDynamics = qds.GetItem(row, query.YDynamics, v => (int)v);
            return cells;
        }

        public static IList<List<ControlCells>> GetCells(IVisio.Page page, IList<int> shapeids)
        {
            var query = new ControlQuery();
            var qds = query.GetFormulasAndResults<double>(page, shapeids);
            var list = new List<List<ControlCells>>(shapeids.Count);
            foreach (var group in qds.Groups)
            {
                var cells_list = new List<ControlCells>(group.Count);
                if (group.Count > 0)
                {
                    for (int i = 0; i < qds.RowCount; i++)
                    {
                        var cells = get_cells_from_row(query, qds, i);
                        cells_list.Add(cells);
                    }
                }
                list.Add(cells_list);
            }
            return list;
        }

        public static IList<ControlCells> GetCells(IVisio.Shape shape)
        {
            var query = new ControlQuery();
            var qds = query.GetFormulasAndResults<double>(shape);
            var cells_list = new List<ControlCells>(qds.RowCount);
            for (int row = 0; row < qds.RowCount; row++)
            {
                var cells = get_cells_from_row(query, qds, row);
                cells_list.Add(cells);
            }
            return cells_list;
        }
    }
}