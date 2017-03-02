namespace VisioAutomation.ShapeSheet.CellGroups
{
    public struct SrcFormulaPair
    {
        public readonly ShapeSheet.Src Src;
        public readonly ShapeSheet.CellValueLiteral Formula;

        public SrcFormulaPair(ShapeSheet.Src src, ShapeSheet.CellValueLiteral formula)
        {
            this.Src = src;
            this.Formula = formula;
        }
    }
}