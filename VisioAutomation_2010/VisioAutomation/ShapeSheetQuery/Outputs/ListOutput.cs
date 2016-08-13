using System.Collections.Generic;

namespace VisioAutomation.ShapeSheetQuery.Outputs
{
    public class ListOutput<T> : IEnumerable<Output<T>>
    {
        private readonly List<Output<T>> _items;

        internal ListOutput()
        {
            this._items = new List<Output<T>>();
        }

        public IEnumerator<Output<T>> GetEnumerator()
        {
            return this._items.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }

        public Output<T> this[int index] => this._items[index];

        internal void Add(Output<T> item)
        {
            this._items.Add(item);
        }

        public int Count => this._items.Count;
    }
}