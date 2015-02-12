using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Query
{
    public class SectionResultList<T> : IEnumerable<SectionResult<T>>
    {
        List<SectionResult<T>> Items;

        internal SectionResultList()
        {
            this.Items = new List<SectionResult<T>>();
        }

        public IEnumerator<SectionResult<T>> GetEnumerator()
        {
            return this.Items.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public SectionResult<T> this[int index]
        {
            get { return this.Items[index]; }
        }


        internal void Add(SectionResult<T> item)
        {
            this.Items.Add(item);
        }

        public int Count
        {
            get { return this.Items.Count; }
        }
    }
}