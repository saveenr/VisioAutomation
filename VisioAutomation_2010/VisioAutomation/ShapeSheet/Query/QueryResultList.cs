using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Query
{
    public class QueryResultList<T> : IEnumerable<QueryResult<T>>
    {
        private readonly List<QueryResult<T>> Items;

        internal QueryResultList()
        {
            this.Items = new List<QueryResult<T>>();
        }

        public IEnumerator<QueryResult<T>> GetEnumerator()
        {
            return this.Items.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }

        public QueryResult<T> this[int index] => this.Items[index];

        internal void Add(QueryResult<T> item)
        {
            this.Items.Add(item);
        }

        public int Count => this.Items.Count;
    }
}