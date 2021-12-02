using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Data
{
    public class CellValueGroup<T> : IEnumerable<CellValueRows<T>>
    {
        // for a given shape, contains rows for every section that was queried
        //
        // {
        //    shapeid
        //    [0] = rows for section0
        //    [1] = rows for section1
        //    [n] = rows for sectionn
        // }

        public readonly int ShapeID;
        private readonly List<CellValueRows<T>> _items;

        internal CellValueGroup(int shapeid, List<CellValueRows<T>> sections)
        {
            this.ShapeID = shapeid;
            this._items = sections;
        }

        public IEnumerator<CellValueRows<T>> GetEnumerator()
        {
            return this._items.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }

        public int Count
        {
            get { return this._items.Count; }
        }

        public CellValueRows<T> this[int index]
        {
            get { return this._items[index]; }
        }
    }
}