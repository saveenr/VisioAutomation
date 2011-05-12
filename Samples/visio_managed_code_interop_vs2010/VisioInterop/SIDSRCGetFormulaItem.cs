namespace VisioInterop
{
    public struct SIDSRCGetFormulaItem
    {
        public short ShapeID;
        public CellSRC CellSRC;

        public SIDSRCGetFormulaItem(short shapeid, CellSRC cellsrc)
        {
            this.ShapeID = shapeid;
            this.CellSRC = cellsrc;
        }
    }
}