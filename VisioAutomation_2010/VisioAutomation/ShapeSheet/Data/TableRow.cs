using VA=VisioAutomation;

namespace VisioAutomation.ShapeSheet.Data
{
    public struct TableRow<T>
    {
        private readonly int _row;

        public int Row
        {
            get { return _row; }
        }

        private Table<T> _table;

        public Table<T> Table
        {
            get { return _table; }
        }

        internal TableRow(Table<T> table,int row)
        {
            this._row = row;
            this._table = table;
        }

        public T this[int column]
        {
            get { return this.Table[this.Row, column]; }
            set { this.Table[this.Row, column] = value; }
        }

        public T this[VA.ShapeSheet.Query.QueryColumn column]
        {
            get { return this.Table[this.Row, column]; }
            set { this.Table[this.Row, column] = value; }
        }

        public int Count
        {
            get { return this.Table.ColumnCount; }
        }
    }
}