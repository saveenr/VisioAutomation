using System.Collections.Generic;
using System.Linq;
using VA=VisioAutomation;

namespace VisioAutomation.ShapeSheet.Streams
{
    public abstract class BaseStream<T>
    {
        protected class ChunkedArray
        {
            public short[] array;
            public int chunksize;

            public ChunkedArray(int capacity, int chunksize)
            {
                this.array = new short[capacity*chunksize];
                this.chunksize = chunksize;
            }

            public int GetIndex(int chunk_index)
            {
                return chunk_index*this.chunksize;
            }

            public void SetItem(int i, short a, short b, short c)
            {
                if (this.chunksize != 3)
                {
                    throw new VA.AutomationException("Incorrect chunksize");
                }

                int pos = this.GetIndex(i);
                this.array[pos + 0] = a;
                this.array[pos + 1] = b;
                this.array[pos + 2] = c;
            }

            public void SetItem(int i, short a, short b, short c, short d)
            {
                if (this.chunksize != 4)
                {
                    throw new VA.AutomationException("Incorrect chunksize");
                }
                int pos = this.GetIndex(i);
                this.array[pos + 0] = a;
                this.array[pos + 1] = b;
                this.array[pos + 2] = c;
                this.array[pos + 3] = d;
            }
        }

        protected readonly ChunkedArray shortarray;
        public readonly int ItemSize;
        private int count;

        protected BaseStream(int capacity, int itemsize)
        {
            this.ItemSize = itemsize;
            this.Capacity = capacity;
            this.count = 0;
            this.shortarray = new ChunkedArray(capacity,itemsize);
        }

        public void Add(T item)
        {
            if (this.IsFull)
            {
                throw new AutomationException("Stream is full");
            }

            int index = this.count;
            this.SetItem(index, item);
            this.count++;
        }

        public int Count
        {
            get { return this.count; }
        }

        public bool IsFull
        {
            get 
            { 
                return this.count >= this.Capacity;
            }
        }

        public short[] Array
        {
            get
            {
                return this.shortarray.array;
            }
        }

        public int Capacity { get; private set; }

        protected void SetItem(int index, T item)
        {
            this.set_item_at_pos(index,item);
        }

        protected abstract void set_item_at_pos(int pos, T item);

        protected void Fill<X>(IList<X> items, System.Func<X, T> get_streamitem)
        {
            foreach (var item in items)
            {
                this.Add(get_streamitem(item));
            }
        }
    }
}