using System.Collections.Generic;

namespace VisioPowerShell.Internal
{
    public class NamedSrcDictionary : NameValueDictionary<VisioAutomation.ShapeSheet.Src>
    {
        public static NamedSrcDictionary FromCells(Models.BaseCells cells)
        {
            var tuples = cells.GetCellTuples();
            return FromCells(tuples);
        }

        public static NamedSrcDictionary FromCells(IEnumerable<Models.CellTuple> tuples)
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