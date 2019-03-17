using System.Collections.Generic;
using System.Linq;

namespace VisioAutomation.ShapeSheet.Query
{
    internal class SectionCacheInfo
    {
        public SectionQuery Query { get; private set; }
        public int RowCount  { get; }

        public short ShapeId;
        internal SectionCacheInfo(SectionQuery sectionquery, int numrows, short shapeid)
        {
            this.Query = sectionquery;
            this.RowCount = numrows;
        }

        public IEnumerable<int> RowIndexes => Enumerable.Range(0, this.RowCount);
    }
}