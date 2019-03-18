using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Query
{
    public class SectionQueryColumns : Columns
    {
        public IVisio.VisSectionIndices SectionIndex { get; private set; }

        internal SectionQueryColumns(IVisio.VisSectionIndices section)
        {
            this.SectionIndex = section;
        }
    }
}