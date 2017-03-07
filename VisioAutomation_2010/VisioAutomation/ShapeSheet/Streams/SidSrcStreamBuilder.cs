
using System.Collections.Generic;

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
            const int sidsrc_length = 4;
            var sidsrcstream = new short[sidsrc_length * sidsrcs.Count];
            for (int i = 0; i < sidsrcs.Count; i++)
            {
                var sidsrc = sidsrcs[i];
                int pos = i * sidsrc_length;
                sidsrcstream[pos + 0] = sidsrc.ShapeID;
                sidsrcstream[pos + 1] = sidsrc.Src.Section;
                sidsrcstream[pos + 2] = sidsrc.Src.Row;
                sidsrcstream[pos + 3] = sidsrc.Src.Cell;
            }
            return sidsrcstream;
        }
    }
}