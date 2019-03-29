using System.Collections.Generic;

namespace VisioPowerShell.Models
{
    public class NameCellDictionary : NameValueDictionary<VisioAutomation.ShapeSheet.Src>
    {
        public static NameCellDictionary FromCells(BaseCells cells)
        {
            var tuples = cells.GetCellTuples();
            return FromCells(tuples);
        }

        public static NameCellDictionary FromCells(IEnumerable<CellTuple> tuples)
        {
            var  dic = new NameCellDictionary();
            foreach (var tuple in tuples)
            {
                dic[tuple.Name] = tuple.Src;
            }
            return dic;
        }
    }
}