using System.Collections.Generic;
using IVisio=Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Data
{
    public class DataRowCollection<T> : IEnumerable<DataRow<T>>
    {
        // Simple list of Rows

        private readonly List<DataRow<T>> _list;

        public readonly int ShapeID = -1;
        public readonly IVisio.VisSectionIndices SectionIndex = IVisio.VisSectionIndices.visSectionInval;

        internal DataRowCollection(int capacity)
        {
            this._list = new List<DataRow<T>>(capacity);
            this.ShapeID = -1;
            this.SectionIndex = IVisio.VisSectionIndices.visSectionInval;
        }

        internal DataRowCollection(int capacity, int shapeid, IVisio.VisSectionIndices section_index)
        {
            this._list = new List<DataRow<T>>(capacity);
            this.ShapeID = shapeid;
            this.SectionIndex = section_index;
        }

        public IEnumerator<DataRow<T>> GetEnumerator()
        {
            return this._list.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        internal void Add(DataRow<T> r)
        {
            this._list.Add(r);
        }

        internal void AddRange(IEnumerable<DataRow<T>> rows)
        {
            this._list.AddRange(rows);
        }

        public int Count => this._list.Count;

        public DataRow<T> this[int index] => this._list[index];
    }
}