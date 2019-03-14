using VASS = VisioAutomation.ShapeSheet;

namespace VisioAutomation.ShapeSheet.CellGroups
{
    public abstract class CellGroupMultiRow : CellGroupBase
    {
        public void SetFormulas(VASS.Writers.SidSrcWriter writer, short shapeid, short row)
        {
            foreach (var pair in this.SidSrcValuePairs_NewRow(shapeid, row))
            {
                writer.SetFormula(pair.ShapeID, pair.Src, pair.Value);
            }
        }

        public void SetFormulas(VASS.Writers.SrcWriter writer, short row)
        {
            foreach (var pair in this.SrcValuePairs_NewRow(row))
            {
                writer.SetFormula(pair.Src, pair.Value);
            }
        }
    }
}