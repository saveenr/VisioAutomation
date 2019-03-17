using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Query
{
    public class RowList<T> : IEnumerable<ShapeCellsRow<T>>
    {

        private readonly List<ShapeCellsRow<T>> _list;

        internal RowList(int capacity)
        {
            this._list = new List<ShapeCellsRow<T>>(capacity);
        }

        public IEnumerator<ShapeCellsRow<T>> GetEnumerator()
        {
            return this._list.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        internal void Add(ShapeCellsRow<T> r)
        {
            this._list.Add(r);
        }

        internal void AddRange(IEnumerable<ShapeCellsRow<T>> rows)
        {
            this._list.AddRange(rows);
        }

        public int Count => this._list.Count;

        public ShapeCellsRow<T> this[int index] => this._list[index];
    }
}