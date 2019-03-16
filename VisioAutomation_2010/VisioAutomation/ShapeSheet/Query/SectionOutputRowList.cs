using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Query
{
    public class SectionOutputRowList<T> : IEnumerable<SectionOutputRow<T>>
    {
        private readonly List<SectionOutputRow<T>> _rows;

        public SectionOutputRowList(int capacity)
        {
            this._rows = new List<SectionOutputRow<T>>(capacity);
        }

        public IEnumerator<SectionOutputRow<T>> GetEnumerator()
        {
            return this._rows.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        internal void Add(SectionOutputRow<T> r)
        {
            this._rows.Add(r);
        }

        public int Count => this._rows.Count;

        public SectionOutputRow<T> this[int index] => this._rows[index];
    }
}