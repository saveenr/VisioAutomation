using SRCCON = VisioAutomation.ShapeSheet.SRCConstants;

namespace VisioAutomation.ShapeSheet.Query.Common
{
    class UserDefinedCellsQuery : CellQuery
    {
        public Query.CellColumn Value { get; set; }
        public Query.CellColumn Prompt { get; set; }

        public UserDefinedCellsQuery()
        {
            var sec = this.AddSection(Microsoft.Office.Interop.Visio.VisSectionIndices.visSectionUser);
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