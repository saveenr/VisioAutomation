namespace VisioAutomation.ShapeSheet.Streams
{
    public class FixedSRCStreamBuilder : FixedStreamBuilder<SRC>
    {
        public FixedSRCStreamBuilder(int capacity) : base(capacity)
        {

        }

        public override int get_chunksize()
        {
            return 3;
        }

        public override void _Add(SRC item)
        {
            this._stream[this._pos++] = item.Section;
            this._stream[this._pos++] = item.Row;
            this._stream[this._pos++] = item.Cell;
        }
    }
}