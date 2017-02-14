namespace VisioAutomation.ShapeSheet
{
    public class SRCStreamBuilder : StreamBuilder<SRC>
    {
        public SRCStreamBuilder() : base()
        {
            
        }

        public SRCStreamBuilder(int capacity) : base(capacity)
        {

        }

        protected override short[] get_stream()
        {
            return SRC.ToStream(this.items);
        }

        internal override void AddSRC(SRC src)
        {
            this.Add(src);
        }
    }
}