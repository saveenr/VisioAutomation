using VisioAutomation.ShapeSheet.Writers;

namespace VisioAutomation.ShapeSheet.CellGroups
{
    public abstract class CellGroupSingleRow : CellGroupBase
    {
        public void SetFormulas(FormulaWriter writer)
        {
            foreach (var pair in this.Pairs)
            {
                writer.SetFormula(pair.SRC, pair.Formula);
            }
        }

        public void SetFormulas(short shapeid, FormulaWriter writer)
        {
            foreach (var pair in this.Pairs)
            {
                writer.SetFormula(shapeid, pair.SRC, pair.Formula);
            }
        }
    }
}