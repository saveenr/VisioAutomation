using IVisio = Microsoft.Office.Interop.Visio;
using VASS = VisioAutomation.ShapeSheet;

namespace VisioAutomation.ShapeSheet.Query
{
    public struct ShapeSectionRowOutput<T>  
    {
        // shapeidn
        // sectionindexn
        // list {
        //     [0] - { cells for (shapeidn,sectionindex0) }
        //     [1] - { cells for (shapeidn,sectionindex1) }
        //     [n] - { cells for (shapeidn,sectionindexn) }
        // }

        public readonly int ShapeID;
        public readonly IVisio.VisSectionIndices SectionIndex;
        public readonly int RowIndex;
        public readonly VASS.Internal.ArraySegment<T> Cells;

        internal ShapeSectionRowOutput(int shapeid, VASS.Internal.ArraySegment<T> cells, IVisio.VisSectionIndices sectionindex, int rowindex)
        {
            this.ShapeID = shapeid;
            this.Cells = cells;
            this.SectionIndex = sectionindex;
            this.RowIndex = rowindex;
        }
    }
}