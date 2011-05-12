namespace VisioInterop
{
    public struct SRCGetResultItem
    {
        public CellSRC CellSRC;
        public short UnitCode;
        
        public SRCGetResultItem(CellSRC cellsrc, short unitcode)
        {
            this.CellSRC = cellsrc;
            this.UnitCode = unitcode;
        }
    }
}