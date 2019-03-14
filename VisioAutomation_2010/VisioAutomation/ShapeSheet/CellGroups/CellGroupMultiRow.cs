using VASS = VisioAutomation.ShapeSheet;

namespace VisioAutomation.ShapeSheet.CellGroups
{
    public abstract class CellGroupMultiRow : CellGroupBase
    {
        public void SetFormulas(VASS.Writers.SrcWriter writer, short row)
        {
            foreach (var pair in this.SrcValuePairs_NewRow(row))
            {
                writer.SetFormula(pair.Src, pair.Value);
            }
        }
    }
}