using VA=VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Controls
{
    public class ControlCells
    {
        public VA.ShapeSheet.CellData<double> X { get; set; }
        public VA.ShapeSheet.CellData<double> Y { get; set; }
        public VA.ShapeSheet.CellData<int> XDynamics { get; set; }
        public VA.ShapeSheet.CellData<int> YDynamics { get; set; }
        public VA.ShapeSheet.CellData<int> CanGlue { get; set; }
        public VA.ShapeSheet.CellData<int> Tip { get; set; }
        public VA.ShapeSheet.CellData<int> XBehavior { get; set; }
        public VA.ShapeSheet.CellData<int> YBehavior { get; set; }

        internal readonly static VA.Controls.ControlQuery query = new VA.Controls.ControlQuery();

        public void Apply(VA.ShapeSheet.Update.SRCUpdate update, short row)
        {
            this._Apply((src, f) => update.SetFormulaIgnoreNull(src, f),row);
        }

        public void Apply(VA.ShapeSheet.Update.SIDSRCUpdate update, short id, short row )
        {
            this._Apply((src, f) => update.SetFormulaIgnoreNull(id, src, f), row);
        }

        internal void _Apply( System.Action<VA.ShapeSheet.SRC,VA.ShapeSheet.FormulaLiteral> func, short row)
        {
            var ctrldef = this;
            func(query.GetCellSRCForRow(query.X, row), ctrldef.X.Formula);
            func(query.GetCellSRCForRow(query.Y, row), ctrldef.Y.Formula);
            func(query.GetCellSRCForRow(query.Glue, row), ctrldef.CanGlue.Formula);
            func(query.GetCellSRCForRow(query.Tip, row), ctrldef.Tip.Formula);
            func(query.GetCellSRCForRow(query.XCon, row), ctrldef.XBehavior.Formula);
            func(query.GetCellSRCForRow(query.YCon, row), ctrldef.YBehavior.Formula);
            func(query.GetCellSRCForRow(query.XDyn, row), ctrldef.XDynamics.Formula);
            func(query.GetCellSRCForRow(query.YDyn, row), ctrldef.YDynamics.Formula);
        }
    }
}