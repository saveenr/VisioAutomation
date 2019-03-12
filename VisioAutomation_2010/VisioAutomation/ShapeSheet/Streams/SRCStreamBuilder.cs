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
            const int src_length = 3;
            var a = new ShapeSheet.Internal.SegmentedArray<short>(this._items.Count, src_length);
            for (int i = 0; i < this._items.Count; i++)
            {
                var src = this._items[i];
                var item = a[i];
                item[0] = src.Section;
                item[1] = src.Row;
                item[2] = src.Cell;
            }
            return new StreamArray(a.Array, Streams.StreamType.Src, this.Count);
        }
    }
}