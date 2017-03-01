using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Query
{
    public class SubQuery
    {
        public string Name { get; private set; }
        public IVisio.VisSectionIndices SectionIndex { get; private set; }
        public SubQueryColumnCollection Columns { get; }
        public int Ordinal { get; }

        internal SubQuery(int ordinal, IVisio.VisSectionIndices section)
        {
            this.Name = VisioAutomation.ShapeSheet.ShapeSheetHelper.GetSectionName(section);
            this.Ordinal = ordinal;
            this.SectionIndex = section;
            this.Columns = new SubQueryColumnCollection();
        }

        public SubQueryColumn AddCell(VisioAutomation.ShapeSheet.SRC src, string name)
        {
            var col = this.Columns.Add(src.Cell, name);
            return col;
        }

        public static implicit operator int(SubQuery col)
        {
            return col.Ordinal;
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

        internal SectionInfo GetSectionInfoForShape(IVisio.Shape shape)
        {
            int rows = this.GetNumRowsForShape(shape);
            var section_info = new SectionInfo(this,rows);
            return section_info;
        }
    }
}