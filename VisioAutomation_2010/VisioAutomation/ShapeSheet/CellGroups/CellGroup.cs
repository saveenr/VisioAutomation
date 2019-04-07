using System.Collections.Generic;
using VASS = VisioAutomation.ShapeSheet;

namespace VisioAutomation.ShapeSheet.CellGroups
{
    public class CellGroup
    {
        public virtual IEnumerable<CellMetadataItem> CellMetadata { get; }

        public IEnumerable<SrcValuePair> GetSrcValuePairs()
        {
            foreach (var pair in this.CellMetadata)
            {
                yield return new SrcValuePair(pair.Src, pair.Value);
            }
        }

        public IEnumerable<SrcValuePair> GetSrcValuePairs_NewRow(short row)
        {
            foreach (var pair in this.GetSrcValuePairs())
            {
                var new_src = pair.Src.CloneWithNewRow(row);
                var new_pair = new SrcValuePair(new_src, pair.Value);
                yield return new_pair;

            }
        }

        public IEnumerable<SidSrcValuePair> GetSidSrcValuePairs_NewRow(short shapeid, short row)
        {
            foreach (var pair in this.GetSrcValuePairs())
            {
                var new_src = pair.Src.CloneWithNewRow(row);
                var new_pair = new SidSrcValuePair(shapeid, new_src, pair.Value);
                yield return new_pair;

            }
        }

        protected CellMetadataItem Create(string name, Src src, CellValue value)
        {
            return new CellMetadataItem(name, src, value.Value);
        }

        internal static System.Func<string,string> row_to_cellgroup(ShapeSheet.Query.Row<string> row, VASS.Query.Columns cols)
        {
            return (s) => row[cols[s].Ordinal];
        }
    }
}