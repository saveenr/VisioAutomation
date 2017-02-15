namespace VisioAutomation.ShapeSheet.CellGroups
{
    public struct SRCFormulaPair
    {
        public readonly ShapeSheet.SRC SRC;
        public readonly ShapeSheet.CellValueLiteral Formula;

        public SRCFormulaPair(ShapeSheet.SRC src, ShapeSheet.CellValueLiteral formula)
        {
            this.SRC = src;
            this.Formula = formula;
        }
    }
}