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
            const int src_length = 4;
            var a = new ShapeSheet.Internal.SegmentedArray<short>(this._items.Count, src_length);
            for (int i = 0; i < this._items.Count; i++)
            {
                var sidsrc = this._items[i];
                var item = a[i];
                item[0] = sidsrc.ShapeID;
                item[1] = sidsrc.Src.Section;
                item[2] = sidsrc.Src.Row;
                item[3] = sidsrc.Src.Cell;
            }

            return new StreamArray(a.Array, Streams.StreamType.SidSrc, this.Count);
        }
    }
}