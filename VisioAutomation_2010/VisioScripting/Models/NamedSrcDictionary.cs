using System;
using System.Collections.Generic;
using System.Linq;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.Query;

namespace VisioScripting.Models
{
    public class NamedSrcDictionary : NameDictionary<Src>
    {
        public ShapeSheetQuery ToQuery(IList<string> Cells)
        {
            var invalid_names = Cells.Where(cellname => !this.ContainsKey(cellname)).ToList();
            if (invalid_names.Count > 0)
            {
                string msg = "Invalid cell names: " + string.Join(",", invalid_names);
                throw new ArgumentException(msg);
            }

            var query = new ShapeSheetQuery();

            foreach (string resolved_cellname in this.ResolveNames(Cells))
            {
                if (!query.Cells.Contains(resolved_cellname))
                {
                    var resolved_src = this[resolved_cellname];
                    query.AddCell(resolved_src, resolved_cellname);
                }
            }
            return query;
        }


        public string[] ExpandCellNames(string [] names)
        {
            // if empty or contains "*" return all the cell names
            if (names == null || names.Length < 1 || names.Contains("*"))
            {
                return this.GetNames().ToArray();
            }

            // otherwise use the names specified
            return names;
        }

        public static NamedSrcDictionary FromCells(BaseCells cells)
        {
            var tuples = cells.GetCellTuples();
            return FromCellTuples(tuples);
        }

        public static NamedSrcDictionary FromCellTuples( IEnumerable<CellTuple> items)
        {
            var  dic = new NamedSrcDictionary();

            foreach (var t in items)
            {
                dic[t.Name] = t.Src;
            }
            return dic;
        }
    }
}