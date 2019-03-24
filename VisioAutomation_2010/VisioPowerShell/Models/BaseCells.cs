using System.Collections.Generic;

namespace VisioPowerShell.Models
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
                    writer.SetValue(id, pair.Src, pair.Formula);
                }
            }
        }

        public static BaseCells CreateCells(CellType celltype)
        {
            if (celltype == VisioPowerShell.Models.CellType.Page)
            {
                return new VisioPowerShell.Models.PageCells();
            }
            else if (celltype == VisioPowerShell.Models.CellType.Shape)
            {
                return new VisioPowerShell.Models.ShapeCells();
            }

            throw new System.ArgumentOutOfRangeException(nameof(celltype));
        }
    }
}