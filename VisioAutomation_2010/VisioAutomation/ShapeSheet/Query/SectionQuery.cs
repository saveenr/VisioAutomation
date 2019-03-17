using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Query
{
    public class SectionQuery
    {
        public ColumnList Columns { get; }
        public IVisio.VisSectionIndices SectionIndex { get; private set; }

        internal SectionQuery(IVisio.VisSectionIndices section)
        {
            this.SectionIndex = section;
            this.Columns = new ColumnList();
        }


        internal short GetNumRowsForShape(IVisio.Shape shape)
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

        internal SectionCacheInfo GetSectionInfoForShape(IVisio.Shape shape)
        {
            int rows = this.GetNumRowsForShape(shape);
            var section_info = new SectionCacheInfo(this,rows, shape.ID16);
            return section_info;
        }
    }
}