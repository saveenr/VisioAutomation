using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Query
{
    public class SectionOutput<T>
    {
        public readonly SectionOutputRowList<T> Rows;
        public readonly IVisio.VisSectionIndices SectionIndex;

        internal SectionOutput(int capacity, IVisio.VisSectionIndices section_index)
        {
            this.Rows = new SectionOutputRowList<T>(capacity);
            this.SectionIndex = section_index;
        }
    }
}