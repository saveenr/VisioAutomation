using System.Collections.Generic;
using IVisio=Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Data
{
    public class CellValueRows<T> : IEnumerable<CellValueRow<T>>
    {
        private readonly List<CellValueRow<T>> _list;

        public readonly int ShapeID = -1;
        public readonly IVisio.VisSectionIndices SectionIndex = IVisio.VisSectionIndices.visSectionInval;

        internal CellValueRows(int capacity)
        {
            this._list = new List<CellValueRow<T>>(capacity);
            this.ShapeID = -1;
            this.SectionIndex = IVisio.VisSectionIndices.visSectionInval;
        }

        internal CellValueRows(int capacity, int shapeid, IVisio.VisSectionIndices section_index)
        {
            this._list = new List<CellValueRow<T>>(capacity);
            this.ShapeID = shapeid;
            this.SectionIndex = section_index;
        }

        public IEnumerator<CellValueRow<T>> GetEnumerator()
        {
            return this._list.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        internal void Add(CellValueRow<T> r)
        {
            this._list.Add(r);
        }

        internal void AddRange(IEnumerable<CellValueRow<T>> rows)
        {
            this._list.AddRange(rows);
        }

        public int Count => this._list.Count;

        public CellValueRow<T> this[int index] => this._list[index];
    }
}