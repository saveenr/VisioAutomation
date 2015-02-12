using VA = VisioAutomation;
using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Query
{
    public class QueryResult<T> : IEnumerable<T>
    {
        public int ShapeID { get; private set; }
        public T[] Cells { get; internal set; }
        public SectionResultList<T> Sections { get; internal set; }

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
            return GetEnumerator();
        }

        public T this[int index]
        {
            get { return this.Cells[index]; }
        }

        public T this[CellColumn col]
        {
            get
            {
                // TODO: Should checking be done on what kind of calolumn it is
                return this.Cells[col.Ordinal];                    
            }
        }

        public int Count
        {
            get { return this.Cells.Length; }
        }
    }
}