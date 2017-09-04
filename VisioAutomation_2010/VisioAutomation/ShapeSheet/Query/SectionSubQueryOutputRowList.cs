using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Query
{
    public class SectionSubQueryOutputRowList<T> : IEnumerable<SectionSubQueryOutputRow<T>>
    {
        private readonly List<SectionSubQueryOutputRow<T>> _rows;

        public SectionSubQueryOutputRowList(int capacity)
        {
            this._rows = new List<SectionSubQueryOutputRow<T>>(capacity);
        }

        public IEnumerator<SectionSubQueryOutputRow<T>> GetEnumerator()
        {
            return this._rows.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        internal void Add(SectionSubQueryOutputRow<T> r)
        {
            this._rows.Add(r);
        }

        public int Count => this._rows.Count;

        public SectionSubQueryOutputRow<T> this[int index] => this._rows[index];
    }
}