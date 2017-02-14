using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet
{
    public abstract class Stream
    {
        public abstract short[] ToStreamArray();
        public abstract int Count();

        public virtual void AddSIDSRC(SIDSRC sidsrc)
        {
            throw new System.NotImplementedException();
        }

        public virtual void AddSRC(SRC src)
        {
            throw new System.NotImplementedException();
        }
    }

    public abstract class Stream<T> : Stream
    {
        protected List<T> items;
        private short[] stream;

        public Stream()
        {
            this.items = new List<T>();
        }

        public override int Count()
        {
            return this.items.Count;
        }

        public Stream(int capacity)
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

        public override short[] ToStreamArray()
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



