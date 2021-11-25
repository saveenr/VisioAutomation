

namespace VisioPowerShell.Models
{
    public abstract class BaseCells
    {
        internal abstract IEnumerable<Internal.CellTuple> EnumCellTuples();

        public void Apply(VASS.Writers.SidSrcWriter writer, short id)
        {
            foreach (var tuple in this.EnumCellTuples())
            {
                if (tuple.Formula != null)
                {
                    writer.SetValue(id, tuple.Src, tuple.Formula);
                }
            }
        }
    }
}