namespace VisioAutomation.ShapeSheet.Streams
{
    public class FixedSrcStreamBuilder : FixedStreamBuilder<Src>
    {
        public FixedSrcStreamBuilder(int capacity) : base(capacity,3)
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