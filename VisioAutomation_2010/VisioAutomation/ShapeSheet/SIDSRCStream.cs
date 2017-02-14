
namespace VisioAutomation.ShapeSheet
{
    public class SIDSRCStream : Stream<SIDSRC>
    {
        public SIDSRCStream() : base()
        {

        }

        public SIDSRCStream(int capacity) : base(capacity)
        {

        }

        protected override short[] get_stream()
        {
            return SIDSRC.ToStream(this.items);
        }

        public override void AddSIDSRC(SIDSRC sidsrc)
        {
            this.Add(sidsrc);
        }
    }
}