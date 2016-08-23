namespace VisioAutomation.ShapeSheet.Queries.Utilities
{
    internal class StreamBuilderSRC: StreamBuilderBase
    {

        public StreamBuilderSRC(int capacity)
            : base(3, capacity)
        {
            
        }

        public void Add(short sec, short row, short cell)
        {
            this.__Add_SRC(sec, row, cell);
        }

        public void Add(SRC cell)
        {
            this.__Add_SRC(cell.Section, cell.Row, cell.Cell);
        }
    }
}