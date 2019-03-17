using System.Collections.Generic;
using System.Linq;

namespace VisioAutomation.ShapeSheet.Query
{
    internal class ShapeCacheItem
    {
        public SectionQuery SectionQuery { get; private set; }

        public int RowCount { get; }

        public short ShapeId { get; }
        internal ShapeCacheItem(SectionQuery sectionquery, int numrows, short shapeid)
        {
            this.SectionQuery = sectionquery;
            this.RowCount = numrows;
            this.ShapeId = shapeid;
        }

        public IEnumerable<int> RowIndexes
        {
            get
            {
                return Enumerable.Range(0, this.RowCount);
            }
        }
    }
}