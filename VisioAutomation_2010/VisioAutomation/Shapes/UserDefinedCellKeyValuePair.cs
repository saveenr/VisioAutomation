namespace VisioAutomation.Shapes
{
    internal class UserDefinedCellKeyValuePair
    {
        public readonly string Name;
        public readonly UserDefinedCellCells Cells;

        public UserDefinedCellKeyValuePair(string name, UserDefinedCellCells cells)
        {
            this.Name = name;
            this.Cells = cells;
        }
    }
}