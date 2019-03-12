using VASS = VisioAutomation.ShapeSheet;

namespace VisioAutomation.ShapeSheet.CellGroups
{
    public abstract class CellGroupSingleRow : CellGroupBase
    {
        public void SetFormulas(VASS.Writers.SrcWriter writer)
        {
            foreach (var pair in this.SrcValuePairs)
            {
                writer.SetFormula(pair.Src, pair.Value);
            }
        }

        public void SetFormulas(VASS.Writers.SidSrcWriter writer, short shapeid)
        {
            foreach (var pair in this.SrcValuePairs)
            {
                writer.SetFormula(shapeid, pair.Src, pair.Value);
            }
        }
    }
}