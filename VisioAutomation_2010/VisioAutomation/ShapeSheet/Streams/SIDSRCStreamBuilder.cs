
namespace VisioAutomation.ShapeSheet.Streams
{
    public class SIDSRCStreamBuilder : StreamBuilder<SIDSRC>
    {
        public SIDSRCStreamBuilder() : base()
        {

        }

        public SIDSRCStreamBuilder(int capacity) : base(capacity)
        {

        }

        protected override short[] build_stream()
        {
            return SIDSRC.ToStream(this.items);
        }
    }
}