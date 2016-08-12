using System.Collections.Generic;

namespace VisioAutomation.ShapeSheetQuery
{
    public struct SectionResultRow<T> 
    {
        public readonly T[] Cells;

        internal SectionResultRow(T[] c)
        {
            this.Cells = c;
        }
    }


    public class SectionSubQueryResult<T>
    {
        public SectionSubQuery Column { get; internal set; }
        public readonly List<SectionResultRow<T>> Rows;

        internal SectionSubQueryResult(int capacity)
        {
            this.Rows = new List<SectionResultRow<T>>(capacity);
        }
    }
}