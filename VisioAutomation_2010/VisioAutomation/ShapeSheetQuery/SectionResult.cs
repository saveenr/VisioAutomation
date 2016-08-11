using System.Collections.Generic;

namespace VisioAutomation.ShapeSheetQuery
{
    public class SectionResult<T> : IEnumerable<T[]>
    {
        public SectionColumn Column { get; internal set; }
        public readonly List<T[]> Rows;

        internal SectionResult(int capacity)
        {
            this.Rows = new List<T[]>(capacity);
        }

        public IEnumerator<T[]> GetEnumerator()
        {
            return this.Rows.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }

        public T[] this[int index] => this.Rows[index];

        internal void Add(T[] item)
        {
            this.Rows.Add(item);
        }

        public int Count => this.Rows.Count;

    }
}