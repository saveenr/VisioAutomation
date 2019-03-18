using System.Collections.Generic;
using System.Linq;

namespace VisioAutomation.ShapeSheet.Query
{
    internal class ShapeCacheItem
    {
        public SectionColumns SectionColumns { get; private set; }

        public int RowCount { get; }

        public short ShapeId { get; }
        internal ShapeCacheItem(SectionColumns sectionquery, int numrows, short shapeid)
        {
            this.SectionColumns = sectionquery;
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