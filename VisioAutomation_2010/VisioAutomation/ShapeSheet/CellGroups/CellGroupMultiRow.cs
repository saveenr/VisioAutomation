using VisioAutomation.ShapeSheet.Writers;

namespace VisioAutomation.ShapeSheet.CellGroups
{
    public abstract class CellGroupMultiRow : CellGroupBase
    {
        public void SetFormulas(SidSrcWriter writer, short shapeid, short row)
        {
            foreach (var pair in this.SrcValuePairs)
            {
                var new_src = pair.Src.CloneWithNewRow(row);
                writer.SetFormula(shapeid, new_src, pair.Value);
            }
        }

        public void SetFormulas(SrcWriter writer, short row)
        {
            foreach (var pair in this.SrcValuePairs)
            {
                var new_src = pair.Src.CloneWithNewRow(row);
                writer.SetFormula(new_src, pair.Value);
            }
        }
    }
}