
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
            var short_array = SidSrc.ToStream(this._items);
            return new StreamArray(short_array,Internal.CellCoord.SidSrc);
        }
    }
}