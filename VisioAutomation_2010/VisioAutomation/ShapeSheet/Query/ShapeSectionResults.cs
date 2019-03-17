using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Query
{
    public class ShapeSectionsResults<T> : IEnumerable<ShapeSectionResult<T>>
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
        private List<ShapeSectionResult<T>> _items;

        internal ShapeSectionsResults(int shape_id, List<ShapeSectionResult<T>> sections) 
        {
            this.ShapeID = shape_id;
            this._items = sections;
        }

        public IEnumerator<ShapeSectionResult<T>> GetEnumerator()
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

        public ShapeSectionResult<T> this[int index]
        {
            get
            {
                return this._items[index];
            }
        }

    }
}