using System.Collections;
using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Query
{
    public class SectionQueryResults<T> : IEnumerable<ShapeSectionsResults<T>>
    {
        // this class contains all the outputs for every shape that was queried
        // think of it this collection as having this shape
        //
        // list {
        //     [0] - { shapeid0, {sections found for shapeid0} }
        //     [1] - { shapeid1, {sections found for shapeid1} }
        //     [n] - { shapeidn, {sections found for shapeidn} }
        // }

        List<ShapeSectionsResults<T>> _list;

        internal SectionQueryResults()
        {
            this._list = new List<ShapeSectionsResults<T>>();
        }

        public void Add(ShapeSectionsResults<T> item)
        {
            this._list.Add(item);
        }

        public IEnumerator<ShapeSectionsResults<T>> GetEnumerator()
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

        public ShapeSectionsResults<T> this[int index]
        {
            get
            {
                return this._list[index];
            }
        }
    }
}