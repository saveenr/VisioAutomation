namespace VisioAutomation.Shapes
{
    internal class CustomPropertyNameCellsPair
    {
        public readonly int ShapeID;
        public readonly int Row;
        public readonly string Name;
        public readonly CustomPropertyCells Cells;


        public CustomPropertyNameCellsPair(int shapeid, int row, string name, CustomPropertyCells cells)
        {
            this.ShapeID = shapeid;
            this.Row = row;
            this.Name = name;
            this.Cells = cells;
        }
    }
}