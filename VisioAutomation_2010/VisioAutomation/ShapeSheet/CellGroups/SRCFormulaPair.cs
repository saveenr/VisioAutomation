namespace VisioAutomation.ShapeSheet.CellGroups
{
    public struct SRCFormulaPair
    {
        public readonly ShapeSheet.SRC SRC;
        public readonly ShapeSheet.FormulaLiteral Formula;

        public SRCFormulaPair(ShapeSheet.SRC src, ShapeSheet.FormulaLiteral formula)
        {
            this.SRC = src;
            this.Formula = formula;
        }
    }
}