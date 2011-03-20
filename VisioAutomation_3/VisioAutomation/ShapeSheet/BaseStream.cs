using System.Collections.Generic;
using System.Linq;

namespace VisioAutomation.ShapeSheet
{
    public abstract class BaseStream<T>
    {
        protected readonly short[] shortarray;
        public readonly int ItemSize;
        private int count;

        protected BaseStream(int capacity, int itemsize)
        {
            this.ItemSize = itemsize;
            this.Capacity = capacity;
            this.count = 0;
            this.shortarray = new short[this.ItemSize * this.Capacity];
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
                return this.shortarray;
            }
        }

        public int Capacity { get; private set; }

        protected void SetItem(int index, T item)
        {
            int pos = this.ItemSize*index;
            this.set_item_at_pos(pos,item);
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