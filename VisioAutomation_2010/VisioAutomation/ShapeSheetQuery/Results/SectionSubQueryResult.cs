using System.Collections.Generic;

namespace VisioAutomation.ShapeSheetQuery.Results
{
    public struct SectionSubQueryResultRow<T> 
    {
        public readonly T[] Cells;

        internal SectionSubQueryResultRow(T[] c)
        {
            this.Cells = c;
        }
    }
    
    public class SectionSubQueryResult<T>
    {
        public SubQuery Column { get; internal set; }
        public readonly List<SectionSubQueryResultRow<T>> Rows;

        internal SectionSubQueryResult(int capacity)
        {
            this.Rows = new List<SectionSubQueryResultRow<T>>(capacity);
        }
    }
}