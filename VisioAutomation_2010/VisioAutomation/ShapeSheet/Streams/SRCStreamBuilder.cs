namespace VisioAutomation.ShapeSheet.Streams
{
    public class SrcStreamBuilder : StreamBuilder<Src>
    {
        public SrcStreamBuilder() : base()
        {
            
        }

        public SrcStreamBuilder(int capacity) : base(capacity)
        {

        }

        protected override short[] build_stream()
        {
            return Src.ToStream(this.items);
        }
    }
}