using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Query
{
    public abstract class StreamBase<T>
    {
        protected List<T> items;
        private short[] stream;

        public StreamBase()
        {
            this.items = new List<T>();
        }

        public int Count => this.items.Count;

        public StreamBase(int capacity)
        {
            this.items = new List<T>(capacity);
        }

        public void Add(T item)
        {
            this.items.Add(item);
            this.stream = null;
        }

        public void AddRange(IEnumerable<T> items)
        {
            this.items.AddRange(items);
            this.stream = null;
        }

        public void Clear()
        {
            this.items.Clear();
            this.stream = null;
        }
        
        public short[] ToStreamArray()
        {
            if (this.stream != null)
            {
                return this.stream;
            }
            else
            {
                this.stream = this.get_stream();
                return this.stream;
            }
        }

        protected abstract short[] get_stream();
    }
}