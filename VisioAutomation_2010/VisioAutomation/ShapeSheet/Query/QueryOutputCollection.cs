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








    public class QueryOutputCollectionCells<T> : IEnumerable<QueryOutputCells<T>>
    {
        private readonly List<QueryOutputCells<T>> items;

        internal QueryOutputCollectionCells()
        {
            this.items = new List<QueryOutputCells<T>>();
        }

        public IEnumerator<QueryOutputCells<T>> GetEnumerator()
        {
            return this.items.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }

        public QueryOutputCells<T> this[int index]
        {
            get { return this.items[index]; }
        }

        internal void Add(QueryOutputCells<T> output)
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

    public class QueryOutputCollectionSections<T> : IEnumerable<QueryOutputSections<T>>
    {
        private readonly List<QueryOutputSections<T>> items;

        internal QueryOutputCollectionSections()
        {
            this.items = new List<QueryOutputSections<T>>();
        }

        public IEnumerator<QueryOutputSections<T>> GetEnumerator()
        {
            return this.items.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }

        public QueryOutputSections<T> this[int index]
        {
            get { return this.items[index]; }
        }

        internal void Add(QueryOutputSections<T> output)
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