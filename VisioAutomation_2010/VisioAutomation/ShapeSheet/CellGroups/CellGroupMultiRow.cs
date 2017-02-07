using VisioAutomation.ShapeSheet.Writers;

namespace VisioAutomation.ShapeSheet.CellGroups
{
    public abstract class CellGroupMultiRow : CellGroupBase
    {
        public void SetFormulas(short shapeid, FormulaWriterSIDSRC writer, short row)
        {
            foreach (var pair in this.Pairs)
            {
                var new_src = pair.SRC.CloneWithNewRow(row);
                writer.SetFormula(shapeid, new_src, pair.Formula);
            }
        }

        public void SetFormulas(FormulaWriterSRC writer, short row)
        {
            foreach (var pair in this.Pairs)
            {
                var new_src = pair.SRC.CloneWithNewRow(row);
                writer.SetFormula(new_src, pair.Formula);
            }
        }

    }
}