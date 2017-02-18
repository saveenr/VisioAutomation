using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Query
{
    public class SubQueryOutputRowCollection<T> : IEnumerable<SubQueryOutputRow<T>>
    {
        private readonly List<SubQueryOutputRow<T>> _rows;

        public SubQueryOutputRowCollection(int capacity)
        {
            this._rows = new List<SubQueryOutputRow<T>>(capacity);
        }

        public IEnumerator<SubQueryOutputRow<T>> GetEnumerator()
        {
            return this._rows.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        internal void Add(SubQueryOutputRow<T> r)
        {
            this._rows.Add(r);
        }

        public int Count => this._rows.Count;

        public SubQueryOutputRow<T> this[int index] => this._rows[index];
    }
}