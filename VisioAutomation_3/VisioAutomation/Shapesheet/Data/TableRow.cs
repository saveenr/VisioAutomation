namespace VisioAutomation.ShapeSheet.Query
{
    public struct TableRow<T>
    {
        private int _index;

        public int Index
        {
            get { return _index; }
        }

        private Table<T> _table;

        public Table<T> Table
        {
            get { return _table; }
        }

        internal TableRow(Table<T> table,int index)
        {
            this._index = index;
            this._table = table;
        }

        public T this[int column]
        {
            get { return this.Table[this.Index, column]; }
            set { this.Table[this.Index, column] = value; }
        }

        public T this[QueryColumn column]
        {
            get { return this.Table[this.Index, column]; }
            set { this.Table[this.Index, column] = value; }
        }

        public int Count
        {
            get { return this.Table.Columns.Count; }
        }
    }
}