using System.Collections.Generic;

namespace VisioAutomation.ShapeSheetQuery.Columns
{
    public class ListColumnBase<T> : IEnumerable<T> where T : ColumnBase
    {
        protected IList<T> _items;
        protected Dictionary<string, T> _dic_columns;

        internal ListColumnBase() : this(0)
        {
        }

        internal ListColumnBase(int capacity)
        {
            this._items = new List<T>(capacity);
            this._dic_columns = new Dictionary<string, T>(capacity);
        }

        public IEnumerator<T> GetEnumerator()
        {
            return (this._items).GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }

        public T this[int index] => this._items[index];

        public T this[string name] => this._dic_columns[name];

        public bool Contains(string name) => this._dic_columns.ContainsKey(name);

        protected string fixup_name(string name)
        {
            if (string.IsNullOrEmpty(name))
            {
                name = string.Format("Col{0}", this._items.Count);
            }
            return name;
        }

        public int Count => this._items.Count;

    }
}