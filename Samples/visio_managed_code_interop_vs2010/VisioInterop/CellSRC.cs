namespace VisioInterop
{
    public struct CellSRC
    {
        public short Section;
        public short Row;
        public short Cell;

        public CellSRC(short section, short row, short cell)
        {
            this.Section = section;
            this.Row = row;
            this.Cell = cell;
        }
    }
}