using VisioAutomation.ShapeSheet.Writers;

namespace VisioAutomation.ShapeSheet.CellGroups
{
    public abstract class CellGroupSingleRow : CellGroupBase
    {
        public void SetFormulas(SrcWriter writer)
        {
            foreach (var pair in this.SrcValuePairs)
            {
                writer.SetFormula(pair.Src, pair.Value);
            }
        }

        public void SetFormulas(SidSrcWriter writer, short shapeid)
        {
            foreach (var pair in this.SrcValuePairs)
            {
                writer.SetFormula(shapeid, pair.Src, pair.Value);
            }
        }
    }
}