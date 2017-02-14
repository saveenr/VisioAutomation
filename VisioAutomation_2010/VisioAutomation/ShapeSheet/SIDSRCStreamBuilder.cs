
namespace VisioAutomation.ShapeSheet
{
    public class SIDSRCShapeSheetStreamBuilder : ShapeSheetStreamBuilder<SIDSRC>
    {
        public SIDSRCShapeSheetStreamBuilder() : base()
        {

        }

        public SIDSRCShapeSheetStreamBuilder(int capacity) : base(capacity)
        {

        }

        protected override ShapeSheetStream build_stream()
        {
            return new SIDSRCStream( SIDSRC.ToStream(this.items) );
        }
        
        internal override void AddSIDSRC(SIDSRC sidsrc)
        {
            this.Add(sidsrc);
        }
    }
}