using System.Collections.Generic;

namespace VisioPowerShell.Models
{
    public class NamedSrcDictionary : NameValueDictionary<VisioAutomation.ShapeSheet.Src>
    {
        public static NamedSrcDictionary FromCells(BaseCells cells)
        {
            var tuples = cells.GetCellTuples();
            return FromCells(tuples);
        }

        public static NamedSrcDictionary FromCells(IEnumerable<CellTuple> tuples)
        {
            var  dic = new NamedSrcDictionary();
            foreach (var tuple in tuples)
            {
                dic[tuple.Name] = tuple.Src;
            }
            return dic;
        }
    }
}