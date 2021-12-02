using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Data
{
    public class RowGroup<T> : IEnumerable<Rows<T>>
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
        private readonly List<Rows<T>> _items;

        internal RowGroup(int shapeid, List<Rows<T>> sections)
        {
            this.ShapeID = shapeid;
            this._items = sections;
        }

        public IEnumerator<Rows<T>> GetEnumerator()
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

        public Rows<T> this[int index]
        {
            get { return this._items[index]; }
        }
    }
}