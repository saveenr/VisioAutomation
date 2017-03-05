using VisioAutomation.ShapeSheet.Internal;

namespace VisioAutomation.ShapeSheet.Streams
{
    public class FixedSrcStreamBuilder : FixedStreamBuilderBase<Src>
    {
        public FixedSrcStreamBuilder(int capacity) : base(capacity, StreamType.Src)
        {

        }

        protected override void _fill_segment_with_item(Utilities.ArraySegment<short> seg, Src item)
        {
            seg[0] = item.Section;
            seg[1] = item.Row;
            seg[2] = item.Cell;
        }
    }
}