using System.Collections.Generic;
using System.Linq;

namespace VisioAutomation.ShapeSheet.Query
{
    internal class SectionInfo
    {
        public SectionQuery Query { get; private set; }
        public int RowCount  { get; }

        internal SectionInfo(SectionQuery sectionquery, int numrows)
        {
            this.Query = sectionquery;
            this.RowCount = numrows;
        }

        public IEnumerable<int> RowIndexes => Enumerable.Range(0, this.RowCount);
    }
}