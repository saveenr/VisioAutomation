using VASS = VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;
using System.Collections.Generic;
using System.Linq;

namespace VisioAutomation.ShapeSheet.Query
{


    public class Row<T> : RowBase<T>
    {
        public readonly IVisio.VisSectionIndices SectionIndex;
        public readonly int RowIndex;

        internal Row(int shapeid, VASS.Internal.ArraySegment<T>  cells) : base(shapeid,cells)
        {
            this.SectionIndex = IVisio.VisSectionIndices.visSectionInval;
            this.RowIndex = -1;
        }

        internal Row(int shapeid, IVisio.VisSectionIndices secindex, int rowindex, VASS.Internal.ArraySegment<T> cells) : base (shapeid,cells)
        {
            this.SectionIndex = secindex;
            this.RowIndex = rowindex;
        }
    }
}