using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet
{
    public abstract class StreamBuilder
    {
        public abstract ShapeSheetStream ToStream();

        public abstract int Count();

        internal virtual void AddSIDSRC(SIDSRC sidsrc)
        {
            throw new System.NotImplementedException();
        }

        internal virtual void AddSRC(SRC src)
        {
            throw new System.NotImplementedException();
        }
    }

    public abstract class StreamBuilder<T> : StreamBuilder
    {
        protected List<T> items;

        public StreamBuilder()
        {
            this.items = new List<T>();
        }

        public override int Count()
        {
            return this.items.Count;
        }

        public StreamBuilder(int capacity)
        {
            this.items = new List<T>(capacity);
        }

        public void Add(T item)
        {
            this.items.Add(item);
        }

        public void AddRange(IEnumerable<T> items)
        {
            this.items.AddRange(items);
        }

        public void Clear()
        {
            this.items.Clear();
        }

        public override ShapeSheetStream ToStream()
        {
             var stream = this.build_stream();
             return stream;
        }

        protected abstract ShapeSheetStream build_stream();
    }
}



