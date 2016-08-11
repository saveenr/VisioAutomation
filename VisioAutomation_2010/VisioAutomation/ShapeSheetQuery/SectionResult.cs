using System.Collections.Generic;

namespace VisioAutomation.ShapeSheetQuery
{
    public class SectionResultRow<T> 
    {
        public readonly T[] Cells;

        internal SectionResultRow(int capacity)
        {
            this.Cells = new T[capacity];
        }

        internal SectionResultRow(T[] c)
        {
            this.Cells = c;
        }
    }


    public class SectionResult<T>
    {
        public SectionColumn Column { get; internal set; }
        public readonly List<SectionResultRow<T>> Rows;

        internal SectionResult(int capacity)
        {
            this.Rows = new List<SectionResultRow<T>>(capacity);
        }
    }
}