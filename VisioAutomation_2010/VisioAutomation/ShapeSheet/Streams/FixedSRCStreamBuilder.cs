namespace VisioAutomation.ShapeSheet.Streams
{
    public class FixedSrcStreamBuilder : FixedStreamBuilder<Src>
    {
        public FixedSrcStreamBuilder(int capacity) : base(capacity)
        {

        }

        public override int get_chunksize()
        {
            return 3;
        }

        protected override void _Add(Src item)
        {
            this._stream[this._pos++] = item.Section;
            this._stream[this._pos++] = item.Row;
            this._stream[this._pos++] = item.Cell;
        }
    }
}