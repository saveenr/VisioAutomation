using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Query
{
    public class ShapeSectionOutput<T>
    {
        public readonly int ShapeID;
        public readonly IVisio.VisSectionIndices SectionIndex;
        public readonly ShapeSectionRowOutputList<T> Rows;

        internal ShapeSectionOutput(int shapeid, int capacity, IVisio.VisSectionIndices section_index)
        {
            this.ShapeID = shapeid;
            this.Rows = new ShapeSectionRowOutputList<T>(shapeid, section_index, capacity);
            this.SectionIndex = section_index;
        }
    }
}