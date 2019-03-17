using System.Collections.Generic;
using System.Linq;

namespace VisioAutomation.ShapeSheet.Query
{
    public class SectionOutputList<T> : IEnumerable<SectionOutput<T>>
    {
        List<SectionOutput<T>> items;

        public SectionOutputList()
        {
            this.items = new List<SectionOutput<T>>();
        }

        public SectionOutputList(int capacity)
        {
            this.items = new List<SectionOutput<T>>(capacity);
        }

        public void Add(SectionOutput<T> item)
        {
            this.items.Add(item);
        }

        public IEnumerator<SectionOutput<T>> GetEnumerator()
        {
            return this.items.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }

        public SectionOutput<T>  this[int index] => this.items[index];

        public int Count
        {
            get
            {
                return this.items.Count;
            }
        }
    }
}