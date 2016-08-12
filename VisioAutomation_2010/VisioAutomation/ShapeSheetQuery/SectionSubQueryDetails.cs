using System.Collections.Generic;
using System.Linq;

namespace VisioAutomation.ShapeSheetQuery
{
    internal class SectionSubQueryDetails
    {
        public SectionSubQuery SectionSubQuery { get; private set; }
        public short ShapeID { get; private set; }
        public int RowCount  { get; }

        internal SectionSubQueryDetails(SectionSubQuery sec_sq, short shapeid, int numrows)
        {
            this.SectionSubQuery = sec_sq;
            this.ShapeID = shapeid;
            this.RowCount = numrows;
        }

        public IEnumerable<int> RowIndexes => Enumerable.Range(0, this.RowCount);
    }
}