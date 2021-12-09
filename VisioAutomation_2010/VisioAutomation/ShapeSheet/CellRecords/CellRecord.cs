using System.Collections.Generic;
using System.Linq;

namespace VisioAutomation.ShapeSheet.CellRecords
{
    public abstract class CellRecord
    {
        public abstract IEnumerable<ColumnMetadata> GetCellMetadata();
        protected ColumnMetadata _create(string name, Core.Src src, Core.CellValue value)
        {
            return new ColumnMetadata(name, src, value.Value);
        }

        internal static System.Func<string, string> getvalueforcol(Data.DataRow<string> row, Data.DataColumns cols)
        {
            return (s) => row[cols[s].Ordinal];
        }
    }
}