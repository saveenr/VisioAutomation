using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Internal
{
    internal class ShapeCacheItem
    {
        public short ShapeId { get; }
        public IVisio.VisSectionIndices SectionIndex { get; }

        public VisioAutomation.ShapeSheet.Query.ColumnCollection ColumnGroup { get; }

        // The RowCount is the data that is being cached
        public int RowCount { get; }

        internal ShapeCacheItem(short shapeid, IVisio.VisSectionIndices sec_index, VisioAutomation.ShapeSheet.Query.ColumnCollection sec_cols,
            int numrows)
        {
            this.ShapeId = shapeid;
            this.SectionIndex = sec_index;
            this.ColumnGroup = sec_cols;
            this.RowCount = numrows;
        }
    }
}