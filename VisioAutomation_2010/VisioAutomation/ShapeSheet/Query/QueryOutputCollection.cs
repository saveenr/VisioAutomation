using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Query
{
    public class QueryOutputCollection<T> : IEnumerable<QueryOutput<T>>
    {
        private readonly List<QueryOutput<T>> items;

        internal QueryOutputCollection()
        {
            this.items = new List<QueryOutput<T>>();
        }

        public IEnumerator<QueryOutput<T>> GetEnumerator()
        {
            return this.items.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }

        public QueryOutput<T> this[int index]
        {
            get { return this.items[index]; }
        }

        internal void Add(QueryOutput<T> output)
        {
            if (output == null)
            {
                throw new System.ArgumentNullException(nameof(output));
                
            }
            this.items.Add(output);
        }

        public int Count
        {
            get { return this.items.Count; }
        }
    }
}