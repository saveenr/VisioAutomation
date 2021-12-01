using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.CellGroups
{
    public class CellGroup
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

        public IEnumerable<SrcValue> GetSrcValuesWithNewRow(short row)
        {
            foreach (var pair in this.GetSrcValues())
            {
                var new_src = pair.Src.CloneWithNewRow(row);
                var new_pair = new SrcValue(new_src, pair.Value);
                yield return new_pair;
            }
        }

        public IEnumerable<SidSrcValue> GetSidSrcValuesWithNewRow(short shapeid, short row)
        {
            foreach (var pair in this.GetSrcValues())
            {
                var new_src = pair.Src.CloneWithNewRow(row);
                var new_pair = new SidSrcValue(shapeid, new_src, pair.Value);
                yield return new_pair;
            }
        }

        protected CellMetadata _create(string name, Core.Src src, Core.CellValue value)
        {
            return new CellMetadata(name, src, value.Value);
        }

        internal static System.Func<string, string> queryrow_to_cellgroup(Query.Row<string> row, Query.Columns cols)
        {
            return (s) => row[cols[s].Ordinal];
        }
    }
}