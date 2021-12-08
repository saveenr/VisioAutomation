using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Data
{
    public class DataRowGroups<T> : IEnumerable<DataRowGroup<T>>
    {
        // Simple list of RowGroups

        private readonly List<DataRowGroup<T>> _list;

        internal DataRowGroups()
        {
            this._list = new List<DataRowGroup<T>>();
        }

        public void Add(DataRowGroup<T> item)
        {
            this._list.Add(item);
        }

        public IEnumerator<DataRowGroup<T>> GetEnumerator()
        {
            return this._list.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }

        public int Count
        {
            get { return this._list.Count; }
        }

        public DataRowGroup<T> this[int index]
        {
            get { return this._list[index]; }
        }
    }
}