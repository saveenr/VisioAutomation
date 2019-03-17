using VASS = VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;
using System.Collections.Generic;
using System.Linq;

namespace VisioAutomation.ShapeSheet.Query
{

    public class ShapeCellsRow<T> : RowBase<T>
    {
        public readonly IVisio.VisSectionIndices SectionIndex;
        public readonly int RowIndex;

        internal ShapeCellsRow(int shapeid, VASS.Internal.ArraySegment<T> cells) : base(shapeid, cells)
        {
            this.SectionIndex = IVisio.VisSectionIndices.visSectionInval;
            this.RowIndex = -1;
        }

        internal ShapeCellsRow(int shapeid, IVisio.VisSectionIndices secindex, int rowindex, VASS.Internal.ArraySegment<T> cells) : base(shapeid, cells)
        {
            this.SectionIndex = secindex;
            this.RowIndex = rowindex;
        }
    }

    public class ShapeSectionCellsRow<T> : RowBase<T>
    {
        public readonly IVisio.VisSectionIndices SectionIndex;
        public readonly int RowIndex;

        internal ShapeSectionCellsRow(int shapeid, VASS.Internal.ArraySegment<T> cells) : base(shapeid, cells)
        {
            this.SectionIndex = IVisio.VisSectionIndices.visSectionInval;
            this.RowIndex = -1;
        }

        internal ShapeSectionCellsRow(int shapeid, IVisio.VisSectionIndices secindex, int rowindex, VASS.Internal.ArraySegment<T> cells) : base(shapeid, cells)
        {
            this.SectionIndex = secindex;
            this.RowIndex = rowindex;
        }
    }
}