namespace VisioAutomation.ShapeSheet.CellGroups
{
    public struct SRCFormulaPair
    {
        public ShapeSheet.SRC SRC;
        public ShapeSheet.FormulaLiteral Formula;

        public SRCFormulaPair(ShapeSheet.SRC src, ShapeSheet.FormulaLiteral formula)
        {
            this.SRC = src;
            this.Formula = formula;
        }
    }
}