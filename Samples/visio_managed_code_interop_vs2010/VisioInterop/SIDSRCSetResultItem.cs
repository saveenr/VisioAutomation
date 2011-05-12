namespace VisioInterop
{
    public struct SIDSRCSetResultItem
    {
        public short ShapeID;
        public CellSRC CellSRC;
        public double Result;
        public short UnitCode;
        
        public SIDSRCSetResultItem(short shapeid, CellSRC cellsrc, double result, short unitcode)
        {
            this.ShapeID = shapeid;
            this.CellSRC = cellsrc;
            this.Result = result;
            this.UnitCode = unitcode;
        }

    }
}