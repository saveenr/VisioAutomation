using System.Collections.Generic;
using System.Collections;

namespace VisioAutomation.Models.DirectedGraph
{
    public class IDList<T> : IEnumerable<T> where T : class
    {
        private readonly Dictionary<string, T> items;

        public IDList()
        {
            this.items = new Dictionary<string, T>();
        }

        public void Add(string id, T item)
        {
            this.items.Add(id, item);
        }

        public T this[string index]
        {
            get { return this.items[index]; }
        }

        public int Count
        {
            get { return this.items.Count; }
        }

        public IEnumerator<T> GetEnumerator()
        {
            foreach (var i in this.items.Values)
            {
                yield return i;
            }
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }

        public bool ContainsID(string id)
        {
            return this.items.ContainsKey(id);
        }

        public IEnumerable<string> IDs
        {
            get
            {
                foreach (var id in this.items.Keys)
                {
                    yield return id;
                }
            }
        }

        public T Find(string id)
        {
            T item = null;
            if (this.items.TryGetValue(id, out item))
            {
                return item;
            }

            return null;
        }
    }

}