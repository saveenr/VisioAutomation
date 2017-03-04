using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Streams
{
    public abstract class StreamBuilder<T>
    {
        protected readonly List<T> _items;

        protected StreamBuilder()
        {
            this._items = new List<T>();
        }

        public int Count => this._items.Count;

        protected StreamBuilder(int capacity)
        {
            this._items = new List<T>(capacity);
        }

        public void Add(T item)
        {
            this._items.Add(item);
        }

        public void AddRange(IEnumerable<T> items)
        {
            this._items.AddRange(items);
        }

        public void Clear()
        {
            this._items.Clear();
        }

        public short[] ToStream()
        {
            var stream = this.build_stream();
            return stream;
        }

        protected abstract short[] build_stream();
    }
}