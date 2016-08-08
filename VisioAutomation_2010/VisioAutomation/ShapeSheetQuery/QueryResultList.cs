using System.Collections.Generic;

namespace VisioAutomation.ShapeSheetQuery
{
    public class QueryResultList<T> : IEnumerable<QueryResult<T>>
    {
        private readonly List<QueryResult<T>> _items;

        internal QueryResultList()
        {
            this._items = new List<QueryResult<T>>();
        }

        public IEnumerator<QueryResult<T>> GetEnumerator()
        {
            return this._items.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }

        public QueryResult<T> this[int index] => this._items[index];

        internal void Add(QueryResult<T> item)
        {
            this._items.Add(item);
        }

        public int Count => this._items.Count;
    }
}