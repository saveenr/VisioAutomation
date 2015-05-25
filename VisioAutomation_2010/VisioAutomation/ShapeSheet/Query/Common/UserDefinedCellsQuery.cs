namespace VisioAutomation.ShapeSheet.Query.Common
{
    class UserDefinedCellsQuery : CellQuery
    {
        public Query.CellColumn Value { get; set; }
        public Query.CellColumn Prompt { get; set; }

        public UserDefinedCellsQuery()
        {
            var sec = this.AddSection(Microsoft.Office.Interop.Visio.VisSectionIndices.visSectionUser);
            this.Value = sec.AddCell(ShapeSheet.SRCConstants.User_Value, "User");
            this.Prompt = sec.AddCell(ShapeSheet.SRCConstants.User_Prompt, "Prompt");
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