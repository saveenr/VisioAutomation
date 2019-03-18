using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Query
{
    public class SectionColumns : Columns
    {
        public IVisio.VisSectionIndices SectionIndex { get; private set; }

        internal SectionColumns(IVisio.VisSectionIndices section)
        {
            this.SectionIndex = section;
        }


        internal short _count_rows_in_visio_shape(IVisio.Shape shape)
        {
            // For visSectionObject we know the result is always going to be 1
            // so avoid making the call tp RowCount[]
            if (this.SectionIndex == IVisio.VisSectionIndices.visSectionObject)
            {
                return 1;
            }

            // For all other cases use RowCount[]
            return shape.RowCount[(short)this.SectionIndex];
        }

        internal ShapeCacheItem _get_shape_cache_item(IVisio.Shape shape)
        {
            int rows = this._count_rows_in_visio_shape(shape);
            var shapecacheitem = new ShapeCacheItem(this, rows, shape.ID16);
            return shapecacheitem;
        }
    }
}