namespace VisioInterop
{
    public struct SIDSRCGetResultItem
    {
        public short ShapeID;
        public CellSRC CellSRC;
        public short UnitCode;

        public SIDSRCGetResultItem(short shapeid, CellSRC cellsrc, short unitcode)
        {
            this.ShapeID = shapeid;
            this.CellSRC = cellsrc;
            this.UnitCode = unitcode;
        }

    }
}