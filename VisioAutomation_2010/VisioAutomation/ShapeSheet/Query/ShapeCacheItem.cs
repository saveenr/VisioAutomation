using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Query
{
    public class ShapeCacheItem
    {

        public short ShapeId { get; }
        public readonly IVisio.VisSectionIndices SectionIndex;

        public SectionQueryColumns SectionColumns { get; private set; }

        // The RowCount is the data that is being cached
        public int RowCount { get; }

        internal ShapeCacheItem(short shapeid, IVisio.VisSectionIndices sec_index, SectionQueryColumns sec_cols, int numrows)
        {
            this.ShapeId = shapeid;
            this.SectionIndex = sec_index;
            this.SectionColumns = sec_cols;
            this.RowCount = numrows;
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