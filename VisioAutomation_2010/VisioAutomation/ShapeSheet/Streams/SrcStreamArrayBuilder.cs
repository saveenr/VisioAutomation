using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Streams
{
    public class SrcStreamArrayBuilder : StreamArrayBuilderBase<Src>
    {
        public SrcStreamArrayBuilder(int capacity) : base(capacity, StreamType.Src)
        {

        }

        protected override void _fill_segment_with_item(ShapeSheet.Internal.ArraySegment<short> seg, Src item)
        {
            seg[0] = item.Section;
            seg[1] = item.Row;
            seg[2] = item.Cell;
        }

        public static VisioAutomation.ShapeSheet.Streams.StreamArray Create(int numcells, IEnumerable<Src> srcs)
        {
            var stream = new VisioAutomation.ShapeSheet.Streams.SrcStreamArrayBuilder(numcells);
            stream.AddRange(srcs);
            return stream.ToStreamArray();
        }
    }
}