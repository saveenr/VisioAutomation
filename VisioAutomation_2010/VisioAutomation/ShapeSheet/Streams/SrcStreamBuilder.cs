using System.Collections.Generic;
using VisioAutomation.Utilities;

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
            var a = new SegmentedArray<short>(srcs.Count, src_length);
            for (int i = 0; i < srcs.Count; i++)
            {
                var src = srcs[i];
                var item = a[i];
                item[0] = src.Section;
                item[1] = src.Row;
                item[2] = src.Cell;
            }
            return a.Array;
        }
    }
}