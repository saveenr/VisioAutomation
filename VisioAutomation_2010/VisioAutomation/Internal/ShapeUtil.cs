using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.Internal
{
    internal class ShapeUtil
    {
        public static void SetSelectGroupFirst(IVisio.Shape shape)
        {
            var src_selectmode = VA.ShapeSheet.SRCConstants.SelectMode;
            var mode = IVisio.VisCellVals.visGrpSelModeGroupOnly;
            var select_mode_cell = shape.CellsSRC[src_selectmode.Section, src_selectmode.Row, src_selectmode.Cell];
            select_mode_cell.FormulaU = ((int)mode).ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
    }
}
