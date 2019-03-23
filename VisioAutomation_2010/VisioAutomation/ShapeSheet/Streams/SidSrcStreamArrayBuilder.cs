using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Streams
{
    public class SidSrcStreamArrayBuilder : StreamArrayBuilderBase<SidSrc>
    {
        public SidSrcStreamArrayBuilder(int capacity) : base(capacity, StreamType.SidSrc)
        {

        }

        protected override void _fill_segment_with_item(ShapeSheet.Internal.ArraySegment<short> seg, SidSrc item)
        {
            seg[0] = item.ShapeID;
            seg[1] = item.Src.Section;
            seg[2] = item.Src.Row;
            seg[3] = item.Src.Cell;
        }  

        public static VisioAutomation.ShapeSheet.Streams.StreamArray Create(int numcells, IEnumerable<SidSrc> sidsrcs)
        {
            var stream = new VisioAutomation.ShapeSheet.Streams.SidSrcStreamArrayBuilder(numcells);
            stream.AddRange(sidsrcs);
            return stream.ToStreamArray();            
        }
    }
}