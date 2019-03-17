using VASS = VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Query
{
    public class Row<T>
    {
        public int ShapeID { get; private set; }
        public readonly IVisio.VisSectionIndices SectionIndex;
        public readonly int RowIndex;
        public readonly VASS.Internal.ArraySegment<T> Cells;

        internal Row(int shapeid, VASS.Internal.ArraySegment<T>  cells)
        {
            this.ShapeID = shapeid;
            this.Cells = cells;
        }

        internal Row(int shapeid, IVisio.VisSectionIndices secindex, int rowindex, VASS.Internal.ArraySegment<T> cells)
        {
            this.ShapeID = shapeid;
            this.SectionIndex = secindex;
            this.RowIndex = rowindex;
            this.Cells = cells;
        }
    }
}