using IVisio = Microsoft.Office.Interop.Visio;
using VASS = VisioAutomation.ShapeSheet;

namespace VisioAutomation.ShapeSheet.Query
{
    public struct SectionOutputRow<T>  
    {
        public readonly VASS.Internal.ArraySegment<T> Cells;
        public readonly int RowIndex;
        public readonly IVisio.VisSectionIndices SectionIndex;

        internal SectionOutputRow(VASS.Internal.ArraySegment<T> cells, IVisio.VisSectionIndices sectionindex, int rowindex)
        {
            this.Cells = cells;
            this.SectionIndex = sectionindex;
            this.RowIndex = rowindex;
        }
    }
}