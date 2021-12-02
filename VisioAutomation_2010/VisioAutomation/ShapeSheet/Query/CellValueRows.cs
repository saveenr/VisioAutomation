using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Query
{
    public class CellValueRows<T> : IEnumerable<CellValueRow<T>>
    {
        private readonly List<CellValueRow<T>> _list;

        internal CellValueRows(int capacity)
        {
            this._list = new List<CellValueRow<T>>(capacity);
        }

        public IEnumerator<CellValueRow<T>> GetEnumerator()
        {
            return this._list.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        internal void Add(CellValueRow<T> r)
        {
            this._list.Add(r);
        }

        internal void AddRange(IEnumerable<CellValueRow<T>> rows)
        {
            this._list.AddRange(rows);
        }

        public int Count => this._list.Count;

        public CellValueRow<T> this[int index] => this._list[index];
    }
}