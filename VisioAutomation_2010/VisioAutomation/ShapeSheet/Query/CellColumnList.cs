using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Query
{
    public class CellColumnList : ColumnListBase
    {
        internal CellColumnList() :
            this(0)
        {
        }

        internal CellColumnList(int capacity) : base(capacity)
        {
        }

        public ColumnBase this[VisioAutomation.ShapeSheet.Src src] => this.dic_src_to_col[src];



    }
}