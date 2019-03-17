using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
namespace VisioAutomation.ShapeSheet.Query
{
    public class ShapeSectionRowOutputList<T> : IEnumerable<ShapeSectionRowOutput<T>>
    {
        private readonly List<ShapeSectionRowOutput<T>> _rows;

        public int ShapeId;
        public readonly IVisio.VisSectionIndices SectionIndex;

        internal ShapeSectionRowOutputList(int shapeid, IVisio.VisSectionIndices secindex, int capacity)
        {
            this.ShapeId = shapeid;
            this.SectionIndex = secindex;
            this._rows = new List<ShapeSectionRowOutput<T>>(capacity);
        }

        public IEnumerator<ShapeSectionRowOutput<T>> GetEnumerator()
        {
            return this._rows.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        internal void Add(ShapeSectionRowOutput<T> r)
        {
            this._rows.Add(r);
        }

        public int Count => this._rows.Count;

        public ShapeSectionRowOutput<T> this[int index] => this._rows[index];
    }
}