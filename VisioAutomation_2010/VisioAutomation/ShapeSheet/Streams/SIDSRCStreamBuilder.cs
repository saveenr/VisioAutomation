
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
            return new StreamArray(SidSrc.ToStream(this._items),Internal.CoordType.SidSrc);
        }
    }
}