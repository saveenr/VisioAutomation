using System.Collections.Generic;
using System.Linq;

namespace VisioAutomation.ShapeSheet.CellRecords
{
    public class CellRecord
    {
        public virtual IEnumerable<CellMetadata> GetCellMetadata()
        {
            throw new System.NotImplementedException();
        }

        protected CellMetadata _create(string name, Core.Src src, Core.CellValue value)
        {
            return new CellMetadata(name, src, value.Value);
        }

        internal static System.Func<string, string> queryrow_to_cellrecord(Data.DataRow<string> row, Data.DataColumns cols)
        {
            return (s) => row[cols[s].Ordinal];
        }
    }
}