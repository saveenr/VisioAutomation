namespace VisioInterop
{
    public struct CellSRC
    {
        public short SectionIndex;
        public short RowIndex;
        public short CellIndex;

        public CellSRC(short section, short row, short cell)
        {
            this.SectionIndex = section;
            this.RowIndex = row;
            this.CellIndex = cell;
        }
    }
}