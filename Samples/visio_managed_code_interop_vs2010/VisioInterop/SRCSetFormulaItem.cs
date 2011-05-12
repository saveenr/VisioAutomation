namespace VisioInterop
{
    public struct SRCSetFormulaItem
    {
        public CellSRC CellSRC;
        public string Formula;
        
        public SRCSetFormulaItem(CellSRC cellsrc, string formula)
        {
            this.CellSRC = cellsrc;
            this.Formula = formula;
        }
    }
}