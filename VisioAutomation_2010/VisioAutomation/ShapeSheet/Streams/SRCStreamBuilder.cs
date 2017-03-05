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

        protected override StreamArray build_stream()
        {
            var short_array = Src.ToStream(this._items);
            return new StreamArray(short_array, Streams.StreamType.Src, this.Count);
        }
    }
}