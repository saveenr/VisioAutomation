using System.Collections.Generic;

namespace VisioAutomation.ShapeSheetQuery
{
    public class QueryResult<T> 
    {
        public int ShapeID { get; private set; }
        public T[] Cells { get; internal set; }
        public List<SectionSubQueryResult<T>> Sections { get; internal set; }

        internal QueryResult(int sid)
        {
            this.ShapeID = sid;
        }
    }
}