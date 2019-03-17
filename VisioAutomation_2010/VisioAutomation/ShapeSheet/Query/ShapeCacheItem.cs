using System.Collections.Generic;
using System.Linq;

namespace VisioAutomation.ShapeSheet.Query
{
    internal class ShapeCacheItem
    {
        public SectionQuery Query { get; private set; }
        public int RowCount  { get; }

        public short ShapeId;
        internal ShapeCacheItem(SectionQuery sectionquery, int numrows, short shapeid)
        {
            this.Query = sectionquery;
            this.RowCount = numrows;
        }

        public IEnumerable<int> RowIndexes => Enumerable.Range(0, this.RowCount);
    }
}