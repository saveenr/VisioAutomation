namespace VisioAutomation.ShapeSheet.CellGroups
{
    public struct SRCFormulaPair
    {
        public readonly ShapeSheet.Src SRC;
        public readonly ShapeSheet.CellValueLiteral Formula;

        public SRCFormulaPair(ShapeSheet.Src src, ShapeSheet.CellValueLiteral formula)
        {
            this.SRC = src;
            this.Formula = formula;
        }
    }
}