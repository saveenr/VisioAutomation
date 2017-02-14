
namespace VisioAutomation.ShapeSheet
{
    public class SidsrcShapeSheetStreamBuilder : ShapeSheetStreamBuilder<SIDSRC>
    {
        public SidsrcShapeSheetStreamBuilder() : base()
        {

        }

        public SidsrcShapeSheetStreamBuilder(int capacity) : base(capacity)
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