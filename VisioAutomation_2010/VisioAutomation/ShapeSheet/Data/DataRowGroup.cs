using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Data
{
    public class DataRowGroup<T> : IEnumerable<DataRows<T>>
    {
        // For a tuple of (shape id, section ids ) contains rows for
        // each pair of (shape and sectionid )
        //
        // Example:
        // rg = RowGroup for a shape 1 and sections A, B, C
        // rg[0] = rows for section A of shape 1
        // rg[1] = rows for section B of shape 1
        // rg[n] = rows for section C of shape 1
        
        public readonly int ShapeID;
        private readonly List<DataRows<T>> _items;

        internal DataRowGroup(int shapeid, List<DataRows<T>> sections)
        {
            this.ShapeID = shapeid;
            this._items = sections;
        }

        public IEnumerator<DataRows<T>> GetEnumerator()
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

        public DataRows<T> this[int index]
        {
            get { return this._items[index]; }
        }
    }
}