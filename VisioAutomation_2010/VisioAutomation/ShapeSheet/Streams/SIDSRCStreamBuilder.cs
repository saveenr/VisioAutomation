
namespace VisioAutomation.ShapeSheet.Streams
{
    public class SIDSRCStreamBuilder : StreamBuilder<SidSrc>
    {
        public SIDSRCStreamBuilder() : base()
        {

        }

        public SIDSRCStreamBuilder(int capacity) : base(capacity)
        {

        }

        protected override short[] build_stream()
        {
            return SidSrc.ToStream(this.items);
        }
    }
}