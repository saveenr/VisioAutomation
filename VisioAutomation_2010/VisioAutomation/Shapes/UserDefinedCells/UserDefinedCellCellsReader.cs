using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet.Query;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Shapes.UserDefinedCells
{
    class UserDefinedCellCellsReader : MultiRowReader<Shapes.UserDefinedCells.UserDefinedCellCells>
    {
        public SubQueryColumn Value { get; set; }
        public SubQueryColumn Prompt { get; set; }

        public UserDefinedCellCellsReader()
        {
            var sec = this.query.AddSubQuery(IVisio.VisSectionIndices.visSectionUser);
            this.Value = sec.AddCell(SrcConstants.UserDelCellValue, nameof(SrcConstants.UserDelCellValue));
            this.Prompt = sec.AddCell(SrcConstants.UserDefCellPrompt, nameof(SrcConstants.UserDefCellPrompt));
        }

        public override Shapes.UserDefinedCells.UserDefinedCellCells CellDataToCellGroup(VisioAutomation.Utilities.ArraySegment<ShapeSheet.CellData> row)
        {
            var cells = new Shapes.UserDefinedCells.UserDefinedCellCells();
            cells.Value = row[this.Value];
            cells.Prompt = row[this.Prompt];
            return cells;
        }
    }
}