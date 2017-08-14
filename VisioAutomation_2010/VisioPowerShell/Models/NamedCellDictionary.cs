using System;
using System.Collections.Generic;
using System.Linq;

namespace VisioPowerShell.Models
{
    public class NamedCellDictionary : NamedDictionary<VisioAutomation.ShapeSheet.Src>
    {
        public VisioAutomation.ShapeSheet.Query.ShapeSheetQuery ToQuery(IList<string> cells)
        {
            var invalid_names = cells.Where(cellname => !this.ContainsKey(cellname)).ToList();

            if (invalid_names.Count > 0)
            {
                string msg = "Invalid cell names: " + string.Join(",", invalid_names);
                throw new ArgumentException(msg);
            }

            var query = new VisioAutomation.ShapeSheet.Query.ShapeSheetQuery();

            foreach (string cell in cells)
            {
                foreach (var resolved_cellname in this.ExpandKeyWildcard(cell))
                {
                    if (!query.Cells.Contains(resolved_cellname))
                    {
                        var resolved_src = this[resolved_cellname];
                        query.AddCell(resolved_src, resolved_cellname);
                    }
                }
            }

            return query;
        }

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