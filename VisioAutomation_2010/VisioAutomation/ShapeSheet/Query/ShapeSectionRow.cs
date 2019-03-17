using IVisio = Microsoft.Office.Interop.Visio;
using VASS = VisioAutomation.ShapeSheet;

namespace VisioAutomation.ShapeSheet.Query
{
    public class ShapeSectionRow<T> : RowBase<T>
    {
        // shapeidn
        // sectionindexn
        // list {
        //     [0] - { cells for (shapeidn,sectionindex0) }
        //     [1] - { cells for (shapeidn,sectionindex1) }
        //     [n] - { cells for (shapeidn,sectionindexn) }
        // }

        public readonly IVisio.VisSectionIndices SectionIndex;
        public readonly int RowIndex;

        internal ShapeSectionRow(int shapeid, IVisio.VisSectionIndices sectionindex, int rowindex, VASS.Internal.ArraySegment<T> cells) : base(shapeid,cells)
        {
            this.SectionIndex = sectionindex;
            this.RowIndex = rowindex;
        }
    }
}