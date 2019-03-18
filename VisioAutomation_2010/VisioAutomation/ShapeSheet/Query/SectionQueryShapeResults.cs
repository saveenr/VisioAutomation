using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Query
{
    public class SectionQueryShapeResults<T> : IEnumerable<SectionQueryShapeRows<T>>
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
        private List<SectionQueryShapeRows<T>> _items;

        internal SectionQueryShapeResults(int shape_id, List<SectionQueryShapeRows<T>> sections) 
        {
            this.ShapeID = shape_id;
            this._items = sections;
        }

        public IEnumerator<SectionQueryShapeRows<T>> GetEnumerator()
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

        public SectionQueryShapeRows<T> this[int index]
        {
            get
            {
                return this._items[index];
            }
        }

    }
}