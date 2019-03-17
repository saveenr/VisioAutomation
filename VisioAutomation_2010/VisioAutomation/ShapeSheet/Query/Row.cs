using VASS = VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;
using System.Collections.Generic;
using System.Linq;

namespace VisioAutomation.ShapeSheet.Query
{
    public class Row<T> : IEnumerable<T>
    {
        public int ShapeID { get; private set; }
        public readonly IVisio.VisSectionIndices SectionIndex;
        public readonly int RowIndex;
        private readonly VASS.Internal.ArraySegment<T> Cells;

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


        public IEnumerator<T> GetEnumerator()
        {
            return this.Cells.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }

        public int Count
        {
            get
            {
                return this.Cells.Count;
            }
        }

        public T this[int index]
        {
            get
            {
                return this.Cells[index];
            }
        }
    }
}