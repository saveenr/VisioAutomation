
using System.Collections.Generic;
using VisioAutomation.Utilities;

namespace VisioAutomation.ShapeSheet.Streams
{
    public class SidSrcStreamBuilder : StreamBuilder<SidSrc>
    {
        public SidSrcStreamBuilder() : base()
        {

        }

        public SidSrcStreamBuilder(int capacity) : base(capacity)
        {

        }

        protected override StreamArray build_stream()
        {
            var short_array = ToStream(this._items);
            return new StreamArray(short_array, Streams.StreamType.SidSrc, this.Count);
        }

        private short[] ToStream(IList<SidSrc> sidsrcs)
        {
            const int src_length = 4;
            var a = new SegmentedArray<short>(sidsrcs.Count, src_length);
            for (int i = 0; i < sidsrcs.Count; i++)
            {
                var sidsrc = sidsrcs[i];
                var item = a[i];
                item[0] = sidsrc.ShapeID;
                item[1] = sidsrc.Src.Section;
                item[2] = sidsrc.Src.Row;
                item[3] = sidsrc.Src.Cell;
            }
            return a.Array;
        }
    }
}