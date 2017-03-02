namespace VisioAutomation.ShapeSheet.Streams
{
    public class FixedSRCStreamBuilder : FixedStreamBuilder<Src>
    {
        public FixedSRCStreamBuilder(int capacity) : base(capacity)
        {

        }

        public override int get_chunksize()
        {
            return 3;
        }

        public override void _Add(Src item)
        {
            this._stream[this._pos++] = item.Section;
            this._stream[this._pos++] = item.Row;
            this._stream[this._pos++] = item.Cell;
        }
    }
}