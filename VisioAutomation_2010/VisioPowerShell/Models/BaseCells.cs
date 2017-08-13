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

        public static BaseCells CreateCells(CellsType type)
        {
            if (type == VisioPowerShell.Models.CellsType.Page)
            {
                return new VisioPowerShell.Models.PageCells();
            }
            else if (type == VisioPowerShell.Models.CellsType.ShapeFormat)
            {
                return new VisioPowerShell.Models.ShapeFormatCells();
            }
            else if (type == VisioPowerShell.Models.CellsType.TextFormat)
            {
                return new VisioPowerShell.Models.TextFormatCells();
            }
            else if (type == VisioPowerShell.Models.CellsType.TextBlock)
            {
                return new VisioPowerShell.Models.TextBlockCells();
            }
            else if (type == VisioPowerShell.Models.CellsType.ShapeXForm)
            {
                return new VisioPowerShell.Models.ShapeXFormCells();
            }
            else if (type == VisioPowerShell.Models.CellsType.Lock)
            {
                return new VisioPowerShell.Models.LockCells();
            }
            else
            {
                throw new System.ArgumentOutOfRangeException();
            }
        }

        public static VisioPowerShell.Models.NamedSrcDictionary GetDictionary(CellsType type)
        {
            var cells = BaseCells.CreateCells(type);
            var dic = VisioPowerShell.Models.NamedSrcDictionary.FromCells(cells);
            return dic;
        }
    }
}