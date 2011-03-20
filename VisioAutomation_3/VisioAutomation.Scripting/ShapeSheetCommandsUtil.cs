using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;
using VA = VisioAutomation;

namespace VisioAutomation.Scripting
{
    public static class ShapeSheetCommandsUtil
    {
        internal static void set_group_select_mode(IVisio.Shape shape, IVisio.VisCellVals mode)
        {
            var src_selectmode = VA.ShapeSheet.SRCConstants.SelectMode;
            var select_mode_cell = shape.CellsSRC[src_selectmode.Section, src_selectmode.Row, src_selectmode.Cell];
            select_mode_cell.FormulaU = ((int)mode).ToString();
        }
    }
}