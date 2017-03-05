using VisioAutomation.ShapeSheet.Writers;

namespace VisioAutomation.ShapeSheet.CellGroups
{
    public abstract class CellGroupMultiRow : CellGroupBase
    {
        public void SetFormulas(short shapeid, SidSrcWriter writer, short row)
        {
            foreach (var pair in this.SrcFormulaPairs)
            {
                var new_src = pair.Src.CloneWithNewRow(row);
                writer.SetFormula(shapeid, new_src, pair.Formula);
            }
        }

        public void SetFormulas(SrcWriter writer, short row)
        {
            foreach (var pair in this.SrcFormulaPairs)
            {
                var new_src = pair.Src.CloneWithNewRow(row);
                writer.SetFormula(new_src, pair.Formula);
            }
        }
    }
}