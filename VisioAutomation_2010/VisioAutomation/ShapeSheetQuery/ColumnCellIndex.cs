namespace VisioAutomation.ShapeSheetQuery
{
    public class ColumnCellIndex : ColumnBase
    {
        public short CellIndex;

        internal ColumnCellIndex(int ordinal, short cell, string name) :
            base(ordinal, name)
        {
            this.CellIndex = cell;
        }
    }
}