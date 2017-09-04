using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Query
{
    public class SectionQueryOutputRowList<T> : IEnumerable<SectionQueryOutputRow<T>>
    {
        private readonly List<SectionQueryOutputRow<T>> _rows;

        public SectionQueryOutputRowList(int capacity)
        {
            this._rows = new List<SectionQueryOutputRow<T>>(capacity);
        }

        public IEnumerator<SectionQueryOutputRow<T>> GetEnumerator()
        {
            return this._rows.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        internal void Add(SectionQueryOutputRow<T> r)
        {
            this._rows.Add(r);
        }

        public int Count => this._rows.Count;

        public SectionQueryOutputRow<T> this[int index] => this._rows[index];
    }
}