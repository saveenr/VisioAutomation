namespace VisioAutomation.ShapeSheet.Queries.Utilities
{
    public class StreamBuilderSIDSRC : StreamBuilderBase
    {

        public StreamBuilderSIDSRC(int capacity) : base(4,capacity)
        {
        }

        public void Add(short shape_id, short sec, short row, short cell)
        {
            this.__Add_SIDSRC(shape_id, sec, row, cell);
        }

        public void Add(short shape_id, SRC cell)
        {
            this.__Add_SIDSRC(shape_id, cell.Section, cell.Row, cell.Cell);
        }
    }
}