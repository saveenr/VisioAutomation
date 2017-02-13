using System.Collections.Generic;
using System.Linq;

namespace VisioAutomation.ShapeSheet.Query
{
    internal class SectionDetails
    {
        public SubQuery SubQuery { get; private set; }
        public int RowCount  { get; }

        internal SectionDetails(SubQuery subquery, int numrows)
        {
            this.SubQuery = subquery;
            this.RowCount = numrows;
        }

        public IEnumerable<int> RowIndexes => Enumerable.Range(0, this.RowCount);
    }
}