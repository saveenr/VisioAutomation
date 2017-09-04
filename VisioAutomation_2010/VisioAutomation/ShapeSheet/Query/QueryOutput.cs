using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Query
{
    public class QueryOutputBase<T> 
    {
        public int ShapeID { get; private set; }

        public int TotalCellCount;

        internal QueryOutputBase(int shape_id)
        {
            this.ShapeID = shape_id;
        }
    }

    public class QueryOutputCells<T>: QueryOutputBase<T>
    {
        public VisioAutomation.Utilities.ArraySegment<T> Cells { get; internal set; }

        internal QueryOutputCells(int shape_id) : base(shape_id)
        {
        }
    }

    public class QueryOutputSections<T> : QueryOutputBase<T>
    {
        public List<SubQueryOutput<T>> Sections { get; internal set; }

        internal QueryOutputSections(int shape_id) : base(shape_id)
        {
        }
    }
}