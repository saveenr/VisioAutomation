namespace VisioAutomation.ShapeSheet.Data
{
    public class TableColumnList<T>
    {
        private readonly Table<T> table;
        private readonly int count;

        internal TableColumnList(Table<T> table, int count)
        {
            this.table = table;
            this.count = count;
        }

        public int Count
        {
            get { return this.count; }
        }
    }
}