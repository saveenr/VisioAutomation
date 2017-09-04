using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Query
{
    public class SubQueryOutput<T>
    {
        public readonly SubQueryOutputRowList<T> Rows;
        public readonly IVisio.VisSectionIndices SectionIndex;

        internal SubQueryOutput(int capacity, IVisio.VisSectionIndices section_index)
        {
            this.Rows = new SubQueryOutputRowList<T>(capacity);
            this.SectionIndex = section_index;
        }
    }
}