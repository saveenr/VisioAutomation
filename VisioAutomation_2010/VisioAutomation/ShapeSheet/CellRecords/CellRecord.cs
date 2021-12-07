using System.Collections.Generic;

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
            foreach (var pair in this.GetCellMetadata())
            {
                yield return new SrcValue(pair.Src, pair.Value);
            }
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