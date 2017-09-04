namespace VisioAutomation.ShapeSheet.Query
{
    public class SectionSubQueryColumn : ColumnBase
    {
        public short CellIndex;

        internal SectionSubQueryColumn(int ordinal, short cell, string name) :
            base(ordinal, name)
        {
            this.CellIndex = cell;
        }
    }
}