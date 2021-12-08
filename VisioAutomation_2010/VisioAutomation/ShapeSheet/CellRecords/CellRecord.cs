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

    public class CellRecords<T> : VisioAutomation.Core.BasicList<T> where T : CellRecord, new()
    {

        public CellRecords() : base()
        {

        }

        public CellRecords(int capacity) : base(capacity)
        {

        }

    }

    public class CellRecordsGroup<T> : VisioAutomation.Core.BasicList<CellRecords<T>> where T : CellRecord, new()
    {

        public CellRecordsGroup() : base()
        {

        }

        public CellRecordsGroup(int capacity) : base(capacity)
        {

        }

    }
}