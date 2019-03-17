using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
namespace VisioAutomation.ShapeSheet.Query
{
    public class ShapeSectionRowOutputList<T> : IEnumerable<ShapeSectionRowOutput<T>>
    {
        // shapeidn
        // list {
        //     [0] - { shapeidn, sectionindex0, {cells for (shapeidn,sectionindex0)} }
        //     [1] - { shapeidn, sectionindex1, {cells for (shapeidn,sectionindex1)} }
        //     [n] - { shapeidn, sectionindexn, {cells for (shapeidn,sectionindexn)} }
        // }

        private readonly List<ShapeSectionRowOutput<T>> _list;

        public int ShapeId;
        public readonly IVisio.VisSectionIndices SectionIndex;

        internal ShapeSectionRowOutputList(int shapeid, IVisio.VisSectionIndices secindex, int capacity)
        {
            this.ShapeId = shapeid;
            this.SectionIndex = secindex;
            this._list = new List<ShapeSectionRowOutput<T>>(capacity);
        }

        public IEnumerator<ShapeSectionRowOutput<T>> GetEnumerator()
        {
            return this._list.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        internal void Add(ShapeSectionRowOutput<T> r)
        {
            this._list.Add(r);
        }

        public int Count => this._list.Count;

        public ShapeSectionRowOutput<T> this[int index] => this._list[index];
    }
}