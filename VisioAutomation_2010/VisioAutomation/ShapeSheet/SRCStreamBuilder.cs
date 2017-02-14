namespace VisioAutomation.ShapeSheet
{
    public class SrcShapeSheetStreamBuilder : ShapeSheetStreamBuilder<SRC>
    {
        public SrcShapeSheetStreamBuilder() : base()
        {
            
        }

        public SrcShapeSheetStreamBuilder(int capacity) : base(capacity)
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