namespace VisioAutomation.ShapeSheet.CellGroups
{
    public struct SRCFormulaPair
    {
        public readonly ShapeSheet.SRC SRC;
        public readonly ShapeSheet.ValueLiteral Formula;

        public SRCFormulaPair(ShapeSheet.SRC src, ShapeSheet.ValueLiteral formula)
        {
            this.SRC = src;
            this.Formula = formula;
        }
    }
}