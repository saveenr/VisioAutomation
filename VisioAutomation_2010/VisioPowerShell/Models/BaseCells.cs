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
                    writer.SetFormula(id, pair.Src, pair.Formula);
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
            else
            {
                throw new System.ArgumentOutOfRangeException(nameof(celltype));
            }
        }

        public static VisioPowerShell.Models.NamedCellDictionary GetDictionary(CellType type)
        {
            var cells = BaseCells.CreateCells(type);
            var dic = VisioPowerShell.Models.NamedCellDictionary.FromCells(cells);
            return dic;
        }
    }
}