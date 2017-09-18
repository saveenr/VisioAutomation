namespace VisioAutomation.ShapeSheet.Streams
{
    public class SrcStreamArrayBuilder : StreamArrayBuilderBase<Src>
    {
        public SrcStreamArrayBuilder(int capacity) : base(capacity, StreamType.Src)
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