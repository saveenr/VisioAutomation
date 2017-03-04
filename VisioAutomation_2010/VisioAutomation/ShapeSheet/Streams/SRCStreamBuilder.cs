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
            return new StreamArray(Src.ToStream(this._items), Internal.CoordType.Src);
        }
    }
}