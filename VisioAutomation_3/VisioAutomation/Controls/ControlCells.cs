using System;
using Microsoft.Office.Interop.Visio;
using VisioAutomation.ShapeSheet;
using VA=VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using System.Collections.Generic;
using System.Linq;

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
            var ctrldef = this;
            update.SetFormulaIgnoreNull(query.GetCellSRCForRow(query.X, row), ctrldef.X.Formula);
            update.SetFormulaIgnoreNull(query.GetCellSRCForRow(query.Y, row), ctrldef.Y.Formula);
            update.SetFormulaIgnoreNull(query.GetCellSRCForRow(query.Glue, row), ctrldef.CanGlue.Formula);
            update.SetFormulaIgnoreNull(query.GetCellSRCForRow(query.Tip, row), ctrldef.Tip.Formula);
            update.SetFormulaIgnoreNull(query.GetCellSRCForRow(query.XCon, row), ctrldef.XBehavior.Formula);
            update.SetFormulaIgnoreNull(query.GetCellSRCForRow(query.YCon, row), ctrldef.YBehavior.Formula);
            update.SetFormulaIgnoreNull(query.GetCellSRCForRow(query.XDyn, row), ctrldef.XDynamics.Formula);
            update.SetFormulaIgnoreNull(query.GetCellSRCForRow(query.YDyn, row), ctrldef.YDynamics.Formula);
        }
    }

}