using System.Collections.Generic;
using System.Linq;

namespace VisioAutomation.ShapeSheet.Query
{
    public class ShapeRowList<T> : IEnumerable<ShapeRow<T>>
    {
        private List<ShapeRow<T>> _list;
        internal ShapeRowList() 
        {
            this._list = new List<ShapeRow<T>>();
        }

        internal ShapeRowList(int capacity)
        {
            this._list = new List<ShapeRow<T>>(capacity);
        }

        public IEnumerator<ShapeRow<T>> GetEnumerator()
        {
            return this._list.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }

        public void AddRange(IEnumerable<ShapeRow<T>> items)
        {
            this._list.AddRange(items);
        }

        public ShapeRow<T> this[int index] => this._list[index];

        public int Count => this._list.Count;
    }
}