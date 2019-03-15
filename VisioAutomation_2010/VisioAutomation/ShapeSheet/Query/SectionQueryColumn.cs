namespace VisioAutomation.ShapeSheet.Query
{
    public class SectionQueryColumn : ColumnBase
    {
        public short CellIndex => this.Src.Cell;

        internal SectionQueryColumn(int ordinal, string name, Src src) :
            base(ordinal, name, src)
        {
        }
    }
}