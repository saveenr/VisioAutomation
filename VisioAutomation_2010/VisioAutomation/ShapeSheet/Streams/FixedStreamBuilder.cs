using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Streams
{
    public abstract class FixedStreamBuilder<T> : StreamBuilderBase
    {
        protected short [] _stream;
        private int _capacity = -1;
        private int _count = 0;
        protected int _pos = 0;

        protected FixedStreamBuilder(int capacity)
        {
            this._capacity = capacity;
            int num_shorts = capacity * this.get_chunksize();
            this._stream = new short[num_shorts];
        }

        public abstract int get_chunksize();

        public override int Count() => this._count;

        public void Add(T item)
        {
            if (this._count >= this._capacity)
            {
                throw new System.ArgumentException("Already full");
            }

            int old_pos = this._pos;
            this._Add(item);
            if (this._pos != old_pos + this.get_chunksize())
            {
                throw new System.ArgumentException();
            }
            this._count++;
        }

        protected abstract void _Add(T item);

        public void AddRange(IEnumerable<T> items)
        {
            foreach (var item in items)
            {
                this.Add(item);
            }
        }

        public override short[] ToStream()
        {
            if (this._count != this._capacity)
            {
                throw new System.ArgumentException("Not full");
            }
            return this._stream;
        }

        public void Clear()
        {
            this._count = 0;
            this._pos = 0;
        }
    }
}