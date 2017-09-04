using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet.Query;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Shapes
{
    class UserDefinedCellCellsReader : ReaderMultiRow<UserDefinedCellCells>
    {
        public SectionQueryColumn Value { get; set; }
        public SectionQueryColumn Prompt { get; set; }

        public UserDefinedCellCellsReader()
        {
            var sec = this.query.SectionQueries.Add(IVisio.VisSectionIndices.visSectionUser);
            this.Value = sec.Columns.Add(SrcConstants.UserDefCellValue, nameof(SrcConstants.UserDefCellValue));
            this.Prompt = sec.Columns.Add(SrcConstants.UserDefCellPrompt, nameof(SrcConstants.UserDefCellPrompt));
        }

        public override UserDefinedCellCells CellDataToCellGroup(VisioAutomation.Utilities.ArraySegment<ShapeSheet.CellData> row)
        {
            var cells = new UserDefinedCellCells();
            cells.Value = row[this.Value];
            cells.Prompt = row[this.Prompt];
            return cells;
        }
    }
}