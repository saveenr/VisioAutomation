using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Query
{
    public class SectionSubQueryOutput<T>
    {
        public readonly SectionSubQueryOutputRowList<T> Rows;
        public readonly IVisio.VisSectionIndices SectionIndex;

        internal SectionSubQueryOutput(int capacity, IVisio.VisSectionIndices section_index)
        {
            this.Rows = new SectionSubQueryOutputRowList<T>(capacity);
            this.SectionIndex = section_index;
        }
    }
}