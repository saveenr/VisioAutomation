using System.Collections.Generic;
using System.Linq;

namespace VisioAutomation.ShapeSheet.Query
{
    public class ShapeRowOutputList<T> : IEnumerable<ShapeRowOutput<T>>
    {
        private List<ShapeRowOutput<T>> _list;
        internal ShapeRowOutputList() 
        {
            this._list = new List<ShapeRowOutput<T>>();
        }

        internal ShapeRowOutputList(int capacity)
        {
            this._list = new List<ShapeRowOutput<T>>(capacity);
        }

        public IEnumerator<ShapeRowOutput<T>> GetEnumerator()
        {
            return this._list.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }

        public void AddRange(IEnumerable<ShapeRowOutput<T>> items)
        {
            this._list.AddRange(items);
        }

        public ShapeRowOutput<T> this[int index] => this._list[index];

        public int Count => this._list.Count;
    }
}