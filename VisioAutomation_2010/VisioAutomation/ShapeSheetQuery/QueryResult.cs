using System.Collections.Generic;

namespace VisioAutomation.ShapeSheetQuery
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
    }
}