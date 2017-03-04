using System.Collections.Generic;
using VisioAutomation.ShapeSheet.Internal;

namespace VisioAutomation.ShapeSheet.Streams
{
    public abstract class FixedStreamBuilderBase<T>
    {
        protected readonly short [] _stream;
        private readonly int _capacity = -1;

        internal readonly CoordType _coordtype;
        protected int _chunksize => this._coordtype == CoordType.SidSrc ? 4 : 3;

        private int _count = 0;
        private int _pos = 0;

        internal FixedStreamBuilderBase(int capacity, CoordType coordtype)
        {
            this._coordtype = coordtype;
            this._capacity = capacity;
            int num_shorts = capacity * this._chunksize;
            this._stream = new short[num_shorts];
        }

        public int Count => this._count;

        public void Add(T item)
        {
            if (this._count >= this._capacity)
            {
                throw new System.ArgumentException("Already full");
            }

            var seg = new Utilities.ArraySegment<short>(this._stream,this._pos,this._chunksize);
            this._fill_segment_with_item(seg,item);
            this._pos = this._pos + this._chunksize;
            this._count++;
        }

        protected abstract void _fill_segment_with_item(Utilities.ArraySegment<short> seg, T item);

        public void AddRange(IEnumerable<T> items)
        {
            foreach (var item in items)
            {
                this.Add(item);
            }
        }

        public VisioAutomation.ShapeSheet.Streams.StreamArray ToStream()
        {
            if (this._count != this._capacity)
            {
                throw new System.ArgumentException("Not full");
            }
            return new StreamArray(this._stream, this._coordtype);
        }

        public void Clear()
        {
            this._count = 0;
            this._pos = 0;
        }
    }
}