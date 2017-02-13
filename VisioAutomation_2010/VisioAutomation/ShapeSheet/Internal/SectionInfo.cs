using System.Collections.Generic;
using System.Linq;
using VisioAutomation.ShapeSheet.Query;

namespace VisioAutomation.ShapeSheet.Internal
{
    internal class SectionInfo
    {
        public SubQuery SubQuery { get; private set; }
        public int RowCount  { get; }

        internal SectionInfo(SubQuery subquery, int numrows)
        {
            this.SubQuery = subquery;
            this.RowCount = numrows;
        }

        public IEnumerable<int> RowIndexes => Enumerable.Range(0, this.RowCount);
    }
}