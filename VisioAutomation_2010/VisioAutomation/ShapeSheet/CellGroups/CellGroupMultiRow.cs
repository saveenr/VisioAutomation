namespace VisioAutomation.ShapeSheet.CellGroups
{
    public abstract class CellGroupMultiRow : CellGroupBase
    {
        public void SetFormulas(short shapeid, ShapeSheetWriterSidSrc writer, short row)
        {
            foreach (var pair in this.SrcFormulaPairs)
            {
                var new_src = pair.Src.CloneWithNewRow(row);
                writer.SetFormula(shapeid, new_src, pair.Formula);
            }
        }

        public void SetFormulas(ShapeSheetWriterSrc writer, short row)
        {
            foreach (var pair in this.SrcFormulaPairs)
            {
                var new_src = pair.Src.CloneWithNewRow(row);
                writer.SetFormula(new_src, pair.Formula);
            }
        }
    }
}