using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Query
{
    public class SectionQueryShapeResults<T> : IEnumerable<SectionShapeRows<T>>
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
        private readonly List<SectionShapeRows<T>> _items;

        internal SectionQueryShapeResults(int shape_id, List<SectionShapeRows<T>> sections) 
        {
            this.ShapeID = shape_id;
            this._items = sections;
        }

        public IEnumerator<SectionShapeRows<T>> GetEnumerator()
        {
            return this._items.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }

        public int Count
        {
            get
            {
                return this._items.Count;
            }
        }

        public SectionShapeRows<T> this[int index]
        {
            get
            {
                return this._items[index];
            }
        }

    }
}