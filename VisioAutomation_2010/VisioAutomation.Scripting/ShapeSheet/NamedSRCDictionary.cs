using System;
using System.Collections.Generic;
using System.Linq;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.Query;

namespace VisioAutomation.Scripting.ShapeSheet
{
    public class NamedSRCDictionary : NameDictionary<SRC>
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
    }
}