using System.Collections.Generic;
using System.Linq;

namespace VisioAutomation.ShapeSheet.Query
{
    public class ShapeCacheItem
    {

        public short ShapeId { get; }
        public SectionColumns SectionColumns { get; private set; }

        // The RowCount is the data that is being cached
        public int RowCount { get; }

        internal ShapeCacheItem(SectionColumns sec_cols, int numrows, short shapeid)
        {
            this.SectionColumns = sec_cols;
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