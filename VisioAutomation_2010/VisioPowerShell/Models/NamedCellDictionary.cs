using System.Collections.Generic;

namespace VisioPowerShell.Models
{
    public class NamedCellDictionary : NamedDictionary<VisioAutomation.ShapeSheet.Src>
    {
        public static NamedCellDictionary FromCells(BaseCells cells)
        {
            var tuples = cells.GetCellTuples();
            return FromCells(tuples);
        }

        public static NamedCellDictionary FromCells(IEnumerable<CellTuple> tuples)
        {
            var  dic = new NamedCellDictionary();
            foreach (var tuple in tuples)
            {
                dic[tuple.Name] = tuple.Src;
            }
            return dic;
        }
    }
}