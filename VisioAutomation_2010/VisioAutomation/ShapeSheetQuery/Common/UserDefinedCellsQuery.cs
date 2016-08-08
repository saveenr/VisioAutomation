using SRCCON = VisioAutomation.ShapeSheet.SRCConstants;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheetQuery.Common
{
    class UserDefinedCellsQuery : CellQuery
    {
        public VisioAutomation.ShapeSheetQuery.CellColumn Value { get; set; }
        public VisioAutomation.ShapeSheetQuery.CellColumn Prompt { get; set; }

        public UserDefinedCellsQuery()
        {
            var sec = this.AddSection(IVisio.VisSectionIndices.visSectionUser);
            this.Value = sec.AddCell(SRCCON.User_Value, nameof(SRCCON.User_Value));
            this.Prompt = sec.AddCell(SRCCON.User_Prompt, nameof(SRCCON.User_Prompt));
        }

        public VisioAutomation.Shapes.UserDefinedCells.UserDefinedCell GetCells(System.Collections.Generic.IList<ShapeSheet.CellData<string>> row)
        {
            var cells = new VisioAutomation.Shapes.UserDefinedCells.UserDefinedCell();
            cells.Value = row[this.Value];
            cells.Prompt = row[this.Prompt];
            return cells;
        }
    }
}