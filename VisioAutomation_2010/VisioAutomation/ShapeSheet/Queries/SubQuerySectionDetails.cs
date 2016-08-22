using System.Collections.Generic;
using System.Linq;

namespace VisioAutomation.ShapeSheet.Queries
{
    internal class SubQuerySectionDetails
    {
        public SubQuery SubQuery { get; private set; }
        public int RowCount  { get; }

        internal SubQuerySectionDetails(SubQuery subquery, int numrows)
        {
            this.SubQuery = subquery;
            this.RowCount = numrows;
        }

        public IEnumerable<int> RowIndexes => Enumerable.Range(0, this.RowCount);
    }
}