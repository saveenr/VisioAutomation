using System.Collections.Generic;

namespace VisioAutomation.Core
{
    public class BasicList<T> : IEnumerable<T>
    {
        private readonly List<T> _list;

        internal BasicList()
        {
            this._list = new List<T>();
        }

        internal BasicList(int capacity)
        {
            this._list = new List<T>(capacity);
        }
        public IEnumerator<T> GetEnumerator()
        {
            return this._list.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public void Add(T r)
        {
            this._list.Add(r);
        }

        public void AddRange(IEnumerable<T> rows)
        {
            this._list.AddRange(rows);
        }

        public int Count
        {
            get
            {
                return this._list.Count;
            }
        }

        public T this[int index]
        {
            get { return this._list[0]; }
            //set { /* set the specified index to value here */ }
        }
    }
}