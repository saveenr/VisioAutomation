using VisioAutomation.ShapeSheet.Writers;

namespace VisioAutomation.ShapeSheet.CellGroups
{
    public abstract class CellGroupSingleRow : CellGroupBase
    {
        public void SetFormulas(ShapeSheetWriter writer)
        {
            foreach (var pair in this.Pairs)
            {
                writer.SetFormula(pair.SRC, pair.Formula);
            }
        }

        public void SetFormulas(short shapeid, ShapeSheetWriter writer)
        {
            foreach (var pair in this.Pairs)
            {
                writer.SetFormula(shapeid, pair.SRC, pair.Formula);
            }
        }
    }
}