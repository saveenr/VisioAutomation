using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Query
{
    public class SectionQueryShapeResults<T> : IEnumerable<SectionShape_CellValueRows<T>>
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
        private readonly List<SectionShape_CellValueRows<T>> _items;

        internal SectionQueryShapeResults(int shapeid, List<SectionShape_CellValueRows<T>> sections)
        {
            this.ShapeID = shapeid;
            this._items = sections;
        }

        public IEnumerator<SectionShape_CellValueRows<T>> GetEnumerator()
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

        public SectionShape_CellValueRows<T> this[int index]
        {
            get { return this._items[index]; }
        }
    }
}