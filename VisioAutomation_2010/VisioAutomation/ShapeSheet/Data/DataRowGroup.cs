using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Data
{
    public class DataRowGroup<T> : IEnumerable<DataRowCollection<T>>
    {
        // For a given shape id
        //   For a given section
        //      contains rows for that shape and section
        //
        // Example:
        // rg = RowGroup for a shape 1 and sections A, B, C
        // rg[0] = rows for section A of shape 1
        // rg[1] = rows for section B of shape 1
        // rg[n] = rows for section C of shape 1
        
        public readonly int ShapeID;
        private readonly List<DataRowCollection<T>> _items;

        internal DataRowGroup(int shapeid, List<DataRowCollection<T>> sections)
        {
            this.ShapeID = shapeid;
            this._items = sections;
        }

        public IEnumerator<DataRowCollection<T>> GetEnumerator()
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

        public DataRowCollection<T> this[int index]
        {
            get { return this._items[index]; }
        }
    }
}