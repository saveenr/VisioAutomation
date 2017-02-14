
namespace VisioAutomation.ShapeSheet
{
    public class SIDSRCStreamBuilder : StreamBuilder<SIDSRC>
    {
        public SIDSRCStreamBuilder() : base()
        {

        }

        public SIDSRCStreamBuilder(int capacity) : base(capacity)
        {

        }

        protected override short[] get_stream()
        {
            return SIDSRC.ToStream(this.items);
        }

        internal override void AddSIDSRC(SIDSRC sidsrc)
        {
            this.Add(sidsrc);
        }
    }
}