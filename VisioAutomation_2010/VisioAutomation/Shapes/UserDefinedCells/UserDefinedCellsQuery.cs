using VisioAutomation.ShapeSheet.CellGroups.Queries;
using VisioAutomation.ShapeSheet.Queries.Columns;
using SRCCON = VisioAutomation.ShapeSheet.SRCConstants;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Shapes.UserDefinedCells
{
    class UserDefinedCellsQuery : CellGroupMultiRowQuery<Shapes.UserDefinedCells.UserDefinedCell>
    {
        public ColumnSubQuery Value { get; set; }
        public ColumnSubQuery Prompt { get; set; }

        public UserDefinedCellsQuery()
        {
            var sec = this.query.AddSubQuery(IVisio.VisSectionIndices.visSectionUser);
            this.Value = sec.AddCell(SRCCON.User_Value, nameof(SRCCON.User_Value));
            this.Prompt = sec.AddCell(SRCCON.User_Prompt, nameof(SRCCON.User_Prompt));
        }

        public override Shapes.UserDefinedCells.UserDefinedCell CellDataToCellGroup(ShapeSheet.CellData[] row)
        {
            var cells = new Shapes.UserDefinedCells.UserDefinedCell();
            cells.Value = row[this.Value];
            cells.Prompt = row[this.Prompt];
            return cells;
        }
    }
}