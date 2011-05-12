namespace VisioInterop
{
    public struct SIDSRCSetFormulaItem
    {
        public short ShapeID;
        public CellSRC CellSRC;
        public string Formula;
        
        public SIDSRCSetFormulaItem(short shapeid, CellSRC cellsrc, string formula)
        {
            this.ShapeID = shapeid;
            this.CellSRC = cellsrc;
            this.Formula = formula;
        }

    }
}