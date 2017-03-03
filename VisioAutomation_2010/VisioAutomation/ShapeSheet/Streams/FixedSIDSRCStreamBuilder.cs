namespace VisioAutomation.ShapeSheet.Streams
{
    public class FixedSidSrcStreamBuilder : FixedStreamBuilder<SidSrc>
    {
        public FixedSidSrcStreamBuilder(int capacity) : base(capacity,4)
        {

        }

        protected override void _Add(Utilities.ArraySegment<short> seg, SidSrc item)
        {
            seg[0] = item.ShapeID;
            seg[1] = item.Section;
            seg[2] = item.Row;
            seg[3] = item.Cell;
        }
    }
}