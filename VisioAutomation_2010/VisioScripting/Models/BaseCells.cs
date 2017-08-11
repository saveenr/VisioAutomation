using System.Collections.Generic;

namespace VisioScripting.Models
{
    public abstract class BaseCells
    {
        public abstract IEnumerable<CellTuple> GetCellTuples();

        public void Apply(VisioAutomation.ShapeSheet.Writers.SidSrcWriter writer, short id)
        {
            foreach (var pair in this.GetCellTuples())
            {
                if (pair.Formula != null)
                {
                    writer.SetFormula(id, pair.Src, pair.Formula);
                }
            }
        }
    }
}