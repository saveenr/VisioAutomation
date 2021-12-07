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

        public IEnumerable<SrcValue> GetSrcValues()
        {
            return this.GetCellMetadata().Select(i => new SrcValue(i.Src, i.Value));
        }

        protected CellMetadata _create(string name, Core.Src src, Core.CellValue value)
        {
            return new CellMetadata(name, src, value.Value);
        }

        internal static System.Func<string, string> queryrow_to_cellrecord(Data.DataRow<string> row, Data.DataColumnCollection cols)
        {
            return (s) => row[cols[s].Ordinal];
        }
    }
}