using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Query
{
    public class QueryResult<T> : IEnumerable<T>
    {
        public int ShapeID { get; private set; }
        public T[] Cells { get; internal set; }
        public List<SectionResult<T>> Sections { get; internal set; }

        internal QueryResult(int sid)
        {
            this.ShapeID = sid;
        }

        public IEnumerator<T> GetEnumerator()
        {
            return ((IList<T>)this.Cells).GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }

        public T this[int index]
        {
            get { return this.Cells[index]; }
        }

        public int Count
        {
            get { return this.Cells.Length; }
        }
    }
}