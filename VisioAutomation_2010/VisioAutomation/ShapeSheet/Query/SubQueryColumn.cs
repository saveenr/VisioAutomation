namespace VisioAutomation.ShapeSheet.Query
{
    public class SubQueryColumn : ColumnBase
    {
        public short CellIndex;

        internal SubQueryColumn(int ordinal, short cell, string name) :
            base(ordinal, name)
        {
            this.CellIndex = cell;
        }
    }
}