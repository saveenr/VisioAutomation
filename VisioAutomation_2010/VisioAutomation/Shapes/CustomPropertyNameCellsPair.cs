
namespace VisioAutomation.Shapes
{

    internal class CustomPropertyNameCellsPair
    {
        public readonly string Name;
        public readonly CustomPropertyCells Cells;

        public CustomPropertyNameCellsPair(string name, CustomPropertyCells cells)
        {
            this.Name = name;
            this.Cells = cells;
        }
    }

}
