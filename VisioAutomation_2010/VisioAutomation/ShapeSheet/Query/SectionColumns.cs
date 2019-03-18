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
    }
}