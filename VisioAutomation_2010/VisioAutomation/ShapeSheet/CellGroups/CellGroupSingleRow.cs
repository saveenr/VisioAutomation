namespace VisioAutomation.ShapeSheet.CellGroups
{
    public abstract class CellGroupSingleRow : CellGroupBase
    {
        public void SetFormulas(ShapeSheetWriter writer)
        {
            foreach (var pair in this.SrcFormulaPairs)
            {
                writer.SetFormula(pair.Src, pair.Formula);
            }
        }

        public void SetFormulas(short shapeid, ShapeSheetWriter writer)
        {
            foreach (var pair in this.SrcFormulaPairs)
            {
                writer.SetFormula(shapeid, pair.Src, pair.Formula);
            }
        }
    }
}