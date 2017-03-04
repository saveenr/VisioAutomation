using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Streams
{
    public abstract class FixedStreamBuilder<T> : StreamBuilderBase
    {
        protected short [] _stream;
        private int _capacity = -1;
        private int _count = 0;
        private int _pos = 0;
        protected int _chunksize;

        protected FixedStreamBuilder(int capacity, int chunksize)
        {
            this._capacity = capacity;
            this._chunksize = chunksize;
            int num_shorts = capacity * this._chunksize;
            this._stream = new short[num_shorts];
        }

        protected override int _GetCount() => this._count;

        public void Add(T item)
        {
            if (this._count >= this._capacity)
            {
                throw new System.ArgumentException("Already full");
            }

            var seg = new Utilities.ArraySegment<short>(this._stream,this._pos,this._chunksize);
            this._Add(seg,item);
            this._pos = this._pos + this._chunksize;
            this._count++;
        }

        protected abstract void _Add(Utilities.ArraySegment<short> seg, T item);

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