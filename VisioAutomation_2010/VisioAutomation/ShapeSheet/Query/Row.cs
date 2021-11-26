using IVisio = Microsoft.Office.Interop.Visio;
using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Query
{

    public class Row<T> : IEnumerable<T>
    {
        public int ShapeID { get; }
        public IVisio.VisSectionIndices SectionIndex { get; }
        private Collections.ArraySegment<T> Cells { get; }

        internal Row(int shapeid, IVisio.VisSectionIndices secindex, Collections.ArraySegment<T> cells)
        {
            this.ShapeID = shapeid;
            this.Cells = cells;
            this.SectionIndex = secindex;
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