using System.Collections.Generic;
using VASS = VisioAutomation.ShapeSheet;

namespace VisioPowerShell.Models
{
    public abstract class BaseCells
    {
        public abstract IEnumerable<CellTuple> GetCellTuples();

        public void Apply(VASS.Writers.SidSrcWriter writer, short id)
        {
            foreach (var tuple in this.GetCellTuples())
            {
                if (tuple.Formula != null)
                {
                    writer.SetValue(id, tuple.Src, tuple.Formula);
                }
            }
        }
    }
}