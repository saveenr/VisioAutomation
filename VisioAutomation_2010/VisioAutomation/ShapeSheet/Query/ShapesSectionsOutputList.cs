using System.Collections;
using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Query
{
    public class ShapesSectionsOutputList<T> : IEnumerable<ShapeSectionOutputList<T>>
    {
        // this class contains all the outputs for every shape that was queried

        List<ShapeSectionOutputList<T>> items;

        internal ShapesSectionsOutputList()
        {
            this.items = new List<ShapeSectionOutputList<T>>();
        }

        public void Add(ShapeSectionOutputList<T> item)
        {
            this.items.Add(item);
        }

        public IEnumerator<ShapeSectionOutputList<T>> GetEnumerator()
        {
            return this.items.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }

        public int Count
        {
            get
            {
                return this.items.Count;
            }
        }

        public ShapeSectionOutputList<T> this[int index]
        {
            get
            {
                return this.items[index];
            }
        }
    }
}