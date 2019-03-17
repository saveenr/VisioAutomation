using System.Collections;
using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Query
{
    public class MultiSectionQueryResults<T> : IEnumerable<ShapeSectionRowsList<T>>
    {
        // this class contains all the outputs for every shape that was queried
        // think of it this collection as having this shape
        //
        // list {
        //     [0] - { shapeid0, {sections found for shapeid0} }
        //     [1] - { shapeid1, {sections found for shapeid1} }
        //     [n] - { shapeidn, {sections found for shapeidn} }
        // }

        List<ShapeSectionRowsList<T>> _list;

        internal MultiSectionQueryResults()
        {
            this._list = new List<ShapeSectionRowsList<T>>();
        }

        public void Add(ShapeSectionRowsList<T> item)
        {
            this._list.Add(item);
        }

        public IEnumerator<ShapeSectionRowsList<T>> GetEnumerator()
        {
            return this._list.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }

        public int Count
        {
            get
            {
                return this._list.Count;
            }
        }

        public ShapeSectionRowsList<T> this[int index]
        {
            get
            {
                return this._list[index];
            }
        }
    }
}