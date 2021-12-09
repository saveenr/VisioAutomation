using System.Collections.Generic;
using System.Linq;

namespace VisioAutomation.ShapeSheet.CellRecords
{
    public abstract class CellRecord
    {
        public abstract IEnumerable<CellMetadata> GetCellMetadata();
        protected CellMetadata _create(string name, Core.Src src, Core.CellValue value)
        {
            return new CellMetadata(name, src, value.Value);
        }

        internal static System.Func<string, string> getvalfromrowfunc(Data.DataRow<string> row, Data.DataColumns cols)
        {
            string func(string s)
            {
                 return row[cols[s].Ordinal];
            }

            return func;
        }
    }
}