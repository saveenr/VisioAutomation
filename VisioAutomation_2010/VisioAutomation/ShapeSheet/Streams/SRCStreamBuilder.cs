namespace VisioAutomation.ShapeSheet.Streams
{
    public class SRCStreamBuilder : StreamBuilder<Src>
    {
        public SRCStreamBuilder() : base()
        {
            
        }

        public SRCStreamBuilder(int capacity) : base(capacity)
        {

        }

        protected override short[] build_stream()
        {
            return Src.ToStream(this.items);
        }
    }
}