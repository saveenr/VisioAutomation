using SRCCON = VisioAutomation.ShapeSheet.SRCConstants;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheetQuery.CommonQueries
{
    class UserDefinedCellsQuery : CellQuery
    {
        public CellColumn Value { get; set; }
        public CellColumn Prompt { get; set; }

        public UserDefinedCellsQuery()
        {
            var sec = this.AddSection(IVisio.VisSectionIndices.visSectionUser);
            this.Value = sec.AddCell(SRCCON.User_Value, nameof(SRCCON.User_Value));
            this.Prompt = sec.AddCell(SRCCON.User_Prompt, nameof(SRCCON.User_Prompt));
        }

        public Shapes.UserDefinedCells.UserDefinedCell GetCells(ShapeSheet.CellData<string>[] row)
        {
            var cells = new Shapes.UserDefinedCells.UserDefinedCell();
            cells.Value = row[this.Value];
            cells.Prompt = row[this.Prompt];
            return cells;
        }
    }
}