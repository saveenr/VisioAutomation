using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Streams
{
    public class SrcStreamBuilder : StreamBuilder<Src>
    {
        public SrcStreamBuilder() : base()
        {
            
        }

        public SrcStreamBuilder(int capacity) : base(capacity)
        {

        }

        protected override StreamArray build_stream()
        {
            var short_array = ToStream(this._items);
            return new StreamArray(short_array, Streams.StreamType.Src, this.Count);
        }

        private short[] ToStream(IList<Src> srcs)
        {
            const int src_length = 3;
            var src_stream = new short[src_length * srcs.Count];
            for (int i = 0; i < srcs.Count; i++)
            {
                var sidsrc = srcs[i];
                int pos = i * src_length;
                src_stream[pos + 0] = sidsrc.Section;
                src_stream[pos + 1] = sidsrc.Row;
                src_stream[pos + 2] = sidsrc.Cell;
            }
            return src_stream;
        }
    }
}