
namespace VisioAutomation.ShapeSheet
{
    public class SIDSRCStreamBuilder : ShapeSheetStreamBuilder<SIDSRC>
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
        
        internal override void AddSIDSRC(SIDSRC sidsrc)
        {
            this.Add(sidsrc);
        }
    }
}