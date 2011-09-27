namespace VisioAutomation.ShapeSheet.Data
{
    public class TableRowList<T>
    {
        private readonly Table<T> table;
        private readonly int count;

        internal TableRowList( Table<T> table, int count)
        {
            this.table = table;
            this.count = count;
        }

        public TableRow<T> this[int row]
        {
            get { return new TableRow<T>(this.table, row); }
        }

        public int Count
        {
            get { return this.count; }
        }
    }
}