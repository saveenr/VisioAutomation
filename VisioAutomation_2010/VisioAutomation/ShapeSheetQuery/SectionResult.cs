using System.Collections.Generic;

namespace VisioAutomation.ShapeSheetQuery
{
    public class SectionResult<T> : IEnumerable<T[]>
    {
        public SectionColumn Column { get; internal set; }
        private readonly List<T[]> _items;

        internal SectionResult(int capacity)
        {
            this._items = new List<T[]>(capacity);
        }

        public IEnumerator<T[]> GetEnumerator()
        {
            return this._items.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }

        public T[] this[int index] => this._items[index];

        internal void Add(T[] item)
        {
            this._items.Add(item);
        }

        public int Count => this._items.Count;
    }
}