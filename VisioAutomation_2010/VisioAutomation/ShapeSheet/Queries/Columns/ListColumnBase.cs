using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Queries.Columns
{
    public class ListColumnBase<TColumn> : IEnumerable<TColumn> where TColumn : ColumnBase
    {
        protected IList<TColumn> _columns;
        protected Dictionary<string, TColumn> _dic_columns;

        internal ListColumnBase() : this(0)
        {
        }

        internal ListColumnBase(int capacity)
        {
            this._columns = new List<TColumn>(capacity);
            this._dic_columns = new Dictionary<string, TColumn>(capacity);
        }

        public IEnumerator<TColumn> GetEnumerator()
        {
            return (this._columns).GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }

        public TColumn this[int index] => this._columns[index];

        public TColumn this[string name] => this._dic_columns[name];

        public bool Contains(string name) => this._dic_columns.ContainsKey(name);

        protected string fixup_name(string name)
        {
            if (string.IsNullOrEmpty(name))
            {
                name = string.Format("Col{0}", this._columns.Count);
            }
            return name;
        }

        public int Count => this._columns.Count;

    }
}