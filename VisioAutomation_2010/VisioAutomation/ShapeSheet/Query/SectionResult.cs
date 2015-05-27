using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Query
{
    public class SectionResult<T> : IEnumerable<T[]>
    {
        public SectionColumn Column { get; internal set; }
        private readonly List<T[]> items;

        internal SectionResult(int capacity)
        {
            this.items = new List<T[]>(capacity);
        }

        public IEnumerator<T[]> GetEnumerator()
        {
            return this.items.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }

        public T[] this[int index] => this.items[index];

        internal void Add(T[] item)
        {
            this.items.Add(item);
        }

        public int Count => this.items.Count;
    }
}