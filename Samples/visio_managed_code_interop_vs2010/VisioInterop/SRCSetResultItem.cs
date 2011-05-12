namespace VisioInterop
{
    public struct SRCSetResultItem
    {
        public CellSRC CellSRC;
        public double Result;
        public short UnitCode;
        
        public SRCSetResultItem(CellSRC cellsrc, double result, short unitcode)
        {
            this.CellSRC = cellsrc;
            this.Result = result;
            this.UnitCode = unitcode;
        }
    }
}