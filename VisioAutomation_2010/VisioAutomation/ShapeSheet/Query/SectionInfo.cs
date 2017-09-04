using System.Collections.Generic;
using System.Linq;

namespace VisioAutomation.ShapeSheet.Query
{
    internal class SectionInfo
    {
        public SectionSubQuery SubQuery { get; private set; }
        public int RowCount  { get; }

        internal SectionInfo(SectionSubQuery subquery, int numrows)
        {
            this.SubQuery = subquery;
            this.RowCount = numrows;
        }

        public IEnumerable<int> RowIndexes => Enumerable.Range(0, this.RowCount);
    }
}