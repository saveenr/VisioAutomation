using System.Collections.Generic;

namespace VisioAutomation.ShapeSheetQuery.Results
{
    public class Result<T> 
    {
        public int ShapeID { get; private set; }
        public T[] Cells { get; internal set; }
        public List<SectionSubQueryResult<T>> Sections { get; internal set; }

        internal Result(int sid)
        {
            this.ShapeID = sid;
        }
    }
}