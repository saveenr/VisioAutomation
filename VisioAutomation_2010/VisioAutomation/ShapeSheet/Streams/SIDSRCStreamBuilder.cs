
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

        protected override short[] build_stream()
        {
            return SidSrc.ToStream(this.items);
        }
    }
}