using System.Collections.Generic;

namespace VisioAutomation.ShapeSheetQuery.Results
{
    public class ListResult<T> : IEnumerable<Result<T>>
    {
        private readonly List<Result<T>> _items;

        internal ListResult()
        {
            this._items = new List<Result<T>>();
        }

        public IEnumerator<Result<T>> GetEnumerator()
        {
            return this._items.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }

        public Result<T> this[int index] => this._items[index];

        internal void Add(Result<T> item)
        {
            this._items.Add(item);
        }

        public int Count => this._items.Count;
    }
}