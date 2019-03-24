using System.Collections.Generic;
using VASS = VisioAutomation.ShapeSheet;

namespace VisioAutomation.ShapeSheet.CellGroups
{
    public class CellGroup
    {
        public IEnumerable<SrcValuePair> SrcValuePairs
        {
            get
            {
                foreach (var pair in this.CellMetadata)
                {
                    yield return new SrcValuePair(pair.Src, pair.Value);
                }
            }
        }
        public virtual IEnumerable<CellMetadataItem> CellMetadata { get; }
        public IEnumerable<SrcValuePair> SrcValuePairs_NewRow(short row)
        {
            foreach (var pair in this.SrcValuePairs)
            {
                var new_src = pair.Src.CloneWithNewRow(row);
                var new_pair = new SrcValuePair(new_src, pair.Value);
                yield return new_pair;

            }
        }

        public IEnumerable<SidSrcValuePair> SidSrcValuePairs_NewRow(short shape_id, short row)
        {
            foreach (var pair in this.SrcValuePairs)
            {
                var new_src = pair.Src.CloneWithNewRow(row);
                var new_pair = new SidSrcValuePair(shape_id, new_src, pair.Value);
                yield return new_pair;

            }
        }

        protected CellMetadataItem Create(string name, Src src, CellValueLiteral value)
        {
            return new CellMetadataItem(name, src, value.Value);
        }

        internal static System.Func<string,string> row_to_cellgroup(ShapeSheet.Query.Row<string> row, VASS.Query.Columns cols)
        {
            return (s) => row[cols[s].Ordinal];
        }
    }
}