using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheetQuery
{
    public class SectionColumn
    {
        public string Name { get; private set; }
        public IVisio.VisSectionIndices SectionIndex { get; private set; }
        public CellColumnList CellColumns { get; }
        public int Ordinal { get; }

        internal SectionColumn(int ordinal, IVisio.VisSectionIndices section)
        {
            this.Name = VisioAutomation.ShapeSheet.ShapeSheetHelper.GetSectionName(section);
            this.Ordinal = ordinal;
            this.SectionIndex = section;
            this.CellColumns = new CellColumnList();
        }

        public CellColumn AddCell(VisioAutomation.ShapeSheet.SRC src, string name)
        {
            var col = this.CellColumns.Add(src, name);
            return col;
        }

        public CellColumn AddCell(short cell, string name)
        {
            var col = this.CellColumns.Add(cell, name);
            return col;
        }

        static public implicit operator int(SectionColumn col)
        {
            return col.Ordinal;
        }
    }
}