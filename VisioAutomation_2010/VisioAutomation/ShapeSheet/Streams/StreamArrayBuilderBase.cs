using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Streams
{
    public abstract class StreamArrayBuilderBase<T>
    {
        private readonly int _capacity = -1;

        protected readonly StreamType _streamtype;
        protected int _chunksize => this._streamtype == StreamType.SidSrc ? 4 : 3;

        private int _count = 0;

        private VisioAutomation.Utilities.SegmentedArray<short> _segarray;

        internal StreamArrayBuilderBase(int capacity, StreamType stream_type)
        {
            this._streamtype = stream_type;
            this._capacity = capacity;
            this._segarray = new VisioAutomation.Utilities.SegmentedArray<short>(capacity,this._chunksize);         
        }

        public int Count => this._count;

        public void Add(T item)
        {
            if (this._count >= this._capacity)
            {
                throw new System.ArgumentException("Already full");
            }

            var seg = this._segarray[this._count];
            this._fill_segment_with_item(seg,item);
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

        public VisioAutomation.ShapeSheet.Streams.StreamArray ToStreamArray()
        {
            if (this._count != this._capacity)
            {
                throw new System.ArgumentException("Not full");
            }
            return new StreamArray(this._segarray.Array, this._streamtype, this._capacity);
        }

        public void Clear()
        {
            this._count = 0;
        }
    }
}