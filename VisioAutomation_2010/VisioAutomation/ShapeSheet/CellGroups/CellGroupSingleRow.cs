namespace VisioAutomation.ShapeSheet.CellGroups
{
    public abstract class CellGroupSingleRow : CellGroupBase
    {
        public void SetFormulas(ShapeSheetWriterSrc writer)
        {
            foreach (var pair in this.SrcFormulaPairs)
            {
                writer.SetFormula(pair.Src, pair.Formula);
            }
        }

        public void SetFormulas(short shapeid, ShapeSheetWriterSidSrc writer)
        {
            foreach (var pair in this.SrcFormulaPairs)
            {
                writer.SetFormula(shapeid, pair.Src, pair.Formula);
            }
        }
    }
}