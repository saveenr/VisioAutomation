using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Query
{
    public class ColumnGroup : ColumnCollection
    {
        public IVisio.VisSectionIndices SectionIndex { get; }

        internal ColumnGroup(IVisio.VisSectionIndices section)
        {
            this.SectionIndex = section;
        }
    }
}