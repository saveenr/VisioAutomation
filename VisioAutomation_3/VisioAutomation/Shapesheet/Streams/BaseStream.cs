using System.Collections.Generic;
using System.Linq;
using VA=VisioAutomation;

namespace VisioAutomation.ShapeSheet.Streams
{
    public abstract class BaseStream<T>
    {
        protected readonly ChunkedArray chunked_array;

        protected class ChunkedArray
        {
            public short[] RawArray { get; private set; }
            public int chunksize;

            public ChunkedArray(int capacity, int chunksize)
            {
                this.RawArray = new short[capacity*chunksize];
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
                this.RawArray[pos + 0] = a;
                this.RawArray[pos + 1] = b;
                this.RawArray[pos + 2] = c;
            }

            public void SetItem(int i, short a, short b, short c, short d)
            {
                if (this.chunksize != 4)
                {
                    throw new VA.AutomationException("Incorrect chunksize");
                }
                int pos = this.GetIndex(i);
                this.RawArray[pos + 0] = a;
                this.RawArray[pos + 1] = b;
                this.RawArray[pos + 2] = c;
                this.RawArray[pos + 3] = d;
            }
        }

        private int count;
        public int ItemSize { get; private set; }

        protected BaseStream(int capacity, int itemsize)
        {
            this.ItemSize = itemsize;
            this.Capacity = capacity;
            this.count = 0;
            this.chunked_array = new ChunkedArray(capacity,itemsize);
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
                return this.chunked_array.RawArray;
            }
        }

        public int Capacity { get; private set; }

        protected abstract void SetItem(int index, T item);

        public void AddRange(IEnumerable<T> items)
        {
            foreach (var item in items)
            {
                this.Add(item);
            }
        }
    }
}