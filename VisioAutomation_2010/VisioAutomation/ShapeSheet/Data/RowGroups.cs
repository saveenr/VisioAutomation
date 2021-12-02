using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Data
{
    public class RowGroups<T> : IEnumerable<RowGroup<T>>
    {
        // this class contains all the outputs for every shape that was queried
        // think of it this collection as having this shape
        //
        // list {
        //     [0] - RowGroup { shapeid0, {sections found for shapeid0} }
        //     [1] - RowGroup { shapeid1, {sections found for shapeid1} }
        //     [n] - RowGroup { shapeidn, {sections found for shapeidn} }
        // }

        private readonly List<RowGroup<T>> _list;

        internal RowGroups()
        {
            this._list = new List<RowGroup<T>>();
        }

        public void Add(RowGroup<T> item)
        {
            this._list.Add(item);
        }

        public IEnumerator<RowGroup<T>> GetEnumerator()
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

        public RowGroup<T> this[int index]
        {
            get { return this._list[index]; }
        }
    }
}