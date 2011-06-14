namespace VisioInterop
{
    public struct SIDSRCGetFormulaItem
    {
        public short ID;
        public CellSRC CellSRC;

        public SIDSRCGetFormulaItem(short shapeid, CellSRC cellsrc)
        {
            this.ID = shapeid;
            this.CellSRC = cellsrc;
        }
    }
}