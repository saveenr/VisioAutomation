using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheetQuery
{
    public class SectionSubQuery
    {
        public string Name { get; private set; }
        public IVisio.VisSectionIndices SectionIndex { get; private set; }
        public SubQueryCellColumnList CellColumns { get; }
        public int Ordinal { get; }

        internal SectionSubQuery(int ordinal, IVisio.VisSectionIndices section)
        {
            this.Name = VisioAutomation.ShapeSheet.ShapeSheetHelper.GetSectionName(section);
            this.Ordinal = ordinal;
            this.SectionIndex = section;
            this.CellColumns = new SubQueryCellColumnList();
        }

        public SubQueryCellColumn AddCell(VisioAutomation.ShapeSheet.SRC src, string name)
        {
            var col = this.CellColumns.Add(src, name);
            return col;
        }

        public SubQueryCellColumn AddCell(short cell, string name)
        {
            var col = this.CellColumns.Add(cell, name);
            return col;
        }

        static public implicit operator int(SectionSubQuery col)
        {
            return col.Ordinal;
        }
    }
}