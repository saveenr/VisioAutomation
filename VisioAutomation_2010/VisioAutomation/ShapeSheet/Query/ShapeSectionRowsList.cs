using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Query
{
    public class ShapeSectionRowsList<T> : IEnumerable<ShapeSectionRows<T>>
    {
        // for a given shape, contains the outputs for every section

        public readonly int ShapeID;
        private List<ShapeSectionRows<T>> _items;

        internal ShapeSectionRowsList(int shape_id, List<ShapeSectionRows<T>> sections) 
        {
            this.ShapeID = shape_id;
            this._items = sections;
        }

        public IEnumerator<ShapeSectionRows<T>> GetEnumerator()
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

        public ShapeSectionRows<T> this[int index]
        {
            get
            {
                return this._items[index];
            }
        }

    }
}