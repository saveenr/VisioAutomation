using System.Collections.Generic;
using System.Linq;

namespace VisioAutomation.ShapeSheet.Query
{
    internal class SectionColumnDetails
    {
        public SectionColumn section_column { get; private set; }
        public short ShapeID { get; private set; }
        public int RowCount  { get; }

        internal SectionColumnDetails(SectionColumn sq, short shapeid, int numrows)
        {
            this.section_column = sq;
            this.ShapeID = shapeid;
            this.RowCount = numrows;
        }

        public IEnumerable<int> RowIndexes
        {
            get { return Enumerable.Range(0, this.RowCount); }
        }
    }
}