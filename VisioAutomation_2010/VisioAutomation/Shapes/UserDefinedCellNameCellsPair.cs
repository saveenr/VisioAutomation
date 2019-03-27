namespace VisioAutomation.Shapes
{
    internal class UserDefinedCellNameCellsPair
    {
        public readonly string Name;
        public readonly UserDefinedCellCells Cells;

        public UserDefinedCellNameCellsPair(string name, UserDefinedCellCells cells)
        {
            this.Name = name;
            this.Cells = cells;
        }
    }
}