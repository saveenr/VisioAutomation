using System.Collections.Generic;

namespace VisioAutomation.ShapeSheetQuery
{
    public class SectionResult<T>
    {
        public SectionColumn Column { get; internal set; }
        public readonly List<T[]> Rows;

        internal SectionResult(int capacity)
        {
            this.Rows = new List<T[]>(capacity);
        }
    }
}