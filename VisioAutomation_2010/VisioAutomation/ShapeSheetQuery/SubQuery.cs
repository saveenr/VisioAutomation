using VisioAutomation.ShapeSheetQuery.Columns;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheetQuery
{
    public class SubQuery
    {
        public string Name { get; private set; }
        public IVisio.VisSectionIndices SectionIndex { get; private set; }
        public ListColumnCellIndex Columns { get; }
        public int Ordinal { get; }

        internal SubQuery(int ordinal, IVisio.VisSectionIndices section)
        {
            this.Name = VisioAutomation.ShapeSheet.ShapeSheetHelper.GetSectionName(section);
            this.Ordinal = ordinal;
            this.SectionIndex = section;
            this.Columns = new ListColumnCellIndex();
        }

        public ColumnCellIndex AddCell(VisioAutomation.ShapeSheet.SRC src, string name)
        {
            var col = this.Columns.Add(src.Cell, name);
            return col;
        }

        public ColumnCellIndex AddCell(short cell, string name)
        {
            var col = this.Columns.Add(cell, name);
            return col;
        }

        public static implicit operator int(SubQuery col)
        {
            return col.Ordinal;
        }
    }
}