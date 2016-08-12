using System.Collections.Generic;
using System.Linq;

namespace VisioAutomation.ShapeSheetQuery
{
    internal class SectionColumnDetails
    {
        public SectionColumn SectionColumn { get; private set; }
        public short ShapeID { get; private set; }
        public int RowCount  { get; }

        internal SectionColumnDetails(SectionColumn sec_col, short shapeid, int numrows)
        {
            this.SectionColumn = sec_col;
            this.ShapeID = shapeid;
            this.RowCount = numrows;
        }

        public IEnumerable<int> RowIndexes => Enumerable.Range(0, this.RowCount);
    }
}