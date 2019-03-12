using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Query
{
    public struct SectionQueryOutputRow<T>  
    {
        public readonly VisioAutomation.ShapeSheet.Internal.ArraySegment<T> Cells;
        public readonly int RowIndex;
        public readonly IVisio.VisSectionIndices SectionIndex;

        internal SectionQueryOutputRow(VisioAutomation.ShapeSheet.Internal.ArraySegment<T> cells, IVisio.VisSectionIndices sectionindex, int rowindex)
        {
            this.Cells = cells;
            this.SectionIndex = sectionindex;
            this.RowIndex = rowindex;
        }
    }
}