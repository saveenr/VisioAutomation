using System.Collections.Generic;
using System.Linq;

namespace VisioAutomation.ShapeSheet.Queries
{
    internal class SubQueryDetails
    {
        public SubQuery SubQuery { get; private set; }
        public short ShapeID { get; private set; }
        public int RowCount  { get; }

        internal SubQueryDetails(SubQuery sec_sq, short shapeid, int numrows)
        {
            this.SubQuery = sec_sq;
            this.ShapeID = shapeid;
            this.RowCount = numrows;
        }

        public IEnumerable<int> RowIndexes => Enumerable.Range(0, this.RowCount);
    }
}