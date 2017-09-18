using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Query
{
    public class SectionQueryOutput<T>
    {
        public readonly SectionQueryOutputRowList<T> Rows;
        public readonly IVisio.VisSectionIndices SectionIndex;

        internal SectionQueryOutput(int capacity, IVisio.VisSectionIndices section_index)
        {
            this.Rows = new SectionQueryOutputRowList<T>(capacity);
            this.SectionIndex = section_index;
        }
    }
}