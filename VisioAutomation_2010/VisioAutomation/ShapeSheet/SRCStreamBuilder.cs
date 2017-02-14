namespace VisioAutomation.ShapeSheet
{
    public class SRCShapeSheetStreamBuilder : ShapeSheetStreamBuilder<SRC>
    {
        public SRCShapeSheetStreamBuilder() : base()
        {
            
        }

        public SRCShapeSheetStreamBuilder(int capacity) : base(capacity)
        {

        }

        protected override ShapeSheetStream build_stream()
        {
            return new SIDSRCStream(SRC.ToStream(this.items));
        }

        internal override void AddSRC(SRC src)
        {
            this.Add(src);
        }
    }
}