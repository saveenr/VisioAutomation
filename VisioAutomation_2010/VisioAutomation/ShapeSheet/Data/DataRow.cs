using IVisio = Microsoft.Office.Interop.Visio;
using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Data
{
    public class DataRow<T> : IEnumerable<T>
    {
        // For a shape id, contains an array of values
        // associated with that shape
        // 
        // SectionIndex = if these cells are known to be part of a section

        public int ShapeID { get; }
        public IVisio.VisSectionIndices SectionIndex { get; }
        private VisioAutomation.Internal.ArraySegment<T> _values { get; }

        internal DataRow(int shapeid, IVisio.VisSectionIndices secindex, VisioAutomation.Internal.ArraySegment<T> values)
        {
            this.ShapeID = shapeid;
            this._values = values;
            this.SectionIndex = secindex;
        }

        public IEnumerator<T> GetEnumerator()
        {
            return this._values.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }

        public int Count
        {
            get { return this._values.Count; }
        }

        public T this[int index]
        {
            get { return this._values[index]; }
        }
    }
}