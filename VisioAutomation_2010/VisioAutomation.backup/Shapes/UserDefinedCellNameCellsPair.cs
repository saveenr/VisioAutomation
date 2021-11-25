namespace VisioAutomation.Shapes
{
    internal class UserDefinedCellNameCellsPair
    {
        public readonly int ShapeID;
        public readonly int Row;
        public readonly string Name;
        public readonly UserDefinedCellCells Cells;
        public UserDefinedCellNameCellsPair(int shapeid, int row, string name, UserDefinedCellCells cells)
        {
            this.ShapeID = shapeid;
            this.Row = row;
            this.Name = name;
            this.Cells = cells;
        }
    }
}