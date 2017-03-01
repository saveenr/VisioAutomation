namespace VisioAutomation.ShapeSheet.Streams
{
    public class FixedSIDSRCStreamBuilder : FixedStreamBuilder<SIDSRC>
    {
        public FixedSIDSRCStreamBuilder(int capacity) : base(capacity)
        {

        }

        public override int get_chunksize()
        {
            return 4;
        }

        public override void _Add(SIDSRC item)
        {
            this._stream[this._pos++] = item.ShapeID;
            this._stream[this._pos++] = item.Section;
            this._stream[this._pos++] = item.Row;
            this._stream[this._pos++] = item.Cell;
        }
    }
}