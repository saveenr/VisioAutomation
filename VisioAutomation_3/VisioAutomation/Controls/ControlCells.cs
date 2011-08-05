using System.Collections.Generic;
using VA=VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Controls
{
    public class ControlCells : VA.ShapeSheet.CellSectionDataGroup
    {
        public VA.ShapeSheet.CellData<int> CanGlue { get; set; }
        public VA.ShapeSheet.CellData<int> Tip { get; set; }
        public VA.ShapeSheet.CellData<double> X { get; set; }
        public VA.ShapeSheet.CellData<double> Y { get; set; }
        public VA.ShapeSheet.CellData<int> YBehavior { get; set; }
        public VA.ShapeSheet.CellData<int> XBehavior { get; set; }
        public VA.ShapeSheet.CellData<int> XDynamics { get; set; }
        public VA.ShapeSheet.CellData<int> YDynamics { get; set; }

        protected override void _Apply(VA.ShapeSheet.CellSectionDataGroup.ApplyFormula func, short row)
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
            cells.CanGlue = qds.GetItem(row, query.CanGlue).Cast(v => (int)v);
            cells.Tip = qds.GetItem(row, query.Tip).Cast(v => (int)v);
            cells.X = qds.GetItem(row, query.X);
            cells.Y = qds.GetItem(row, query.Y);
            cells.YBehavior = qds.GetItem(row, query.YBehavior).Cast(v => (int)v);
            cells.XBehavior = qds.GetItem(row, query.XBehavior).Cast(v => (int)v);
            cells.XDynamics = qds.GetItem(row, query.XDynamics).Cast(v => (int)v);
            cells.YDynamics = qds.GetItem(row, query.YDynamics).Cast(v => (int)v);
            return cells;
        }

        internal static IList<List<ControlCells>> GetCells(IVisio.Page page, IList<int> shapeids)
        {
            var query = new ControlQuery();
            return VA.ShapeSheet.CellSectionDataGroup._GetCells(page, shapeids, query, get_cells_from_row);
        }

        internal static IList<ControlCells> GetCells(IVisio.Shape shape)
        {
            var query = new ControlQuery();
            return VA.ShapeSheet.CellSectionDataGroup._GetCells(shape, query, get_cells_from_row);
        }
    }
}