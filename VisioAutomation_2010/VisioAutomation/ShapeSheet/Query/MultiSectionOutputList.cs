using System.Collections;
using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Query
{
    public class MultiSectionOuputList<T> : IEnumerable<MultiSectionOutput<T>>
    {
        List<MultiSectionOutput<T>> items;

        internal MultiSectionOuputList()
        {
            this.items = new List<MultiSectionOutput<T>>();
        }

        public void Add(MultiSectionOutput<T> item)
        {
            this.items.Add(item);
        }

        public IEnumerator<MultiSectionOutput<T>> GetEnumerator()
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

        public MultiSectionOutput<T> this[int index]
        {
            get
            {
                return this.items[index];
            }
        }
    }
}