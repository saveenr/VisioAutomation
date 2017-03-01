namespace VisioAutomation.ShapeSheet.Streams
{
    public class SRCStreamBuilder : StreamBuilder<SRC>
    {
        public SRCStreamBuilder() : base()
        {
            
        }

        public SRCStreamBuilder(int capacity) : base(capacity)
        {

        }

        protected override short[] build_stream()
        {
            return SRC.ToStream(this.items);
        }
    }
}