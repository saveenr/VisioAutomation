using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Query
{
    public class QueryOutputBase<T> 
    {
        public int ShapeID { get; private set; }

        internal readonly int __totalcellcount;

        internal QueryOutputBase(int shape_id, int count)
        {
            this.ShapeID = shape_id;
            this.__totalcellcount = count;
        }
    }

    public class QueryOutputCells<T>: QueryOutputBase<T>
    {
        public VisioAutomation.Utilities.ArraySegment<T> Cells { get; internal set; }

        internal QueryOutputCells(int shape_id, int count, VisioAutomation.Utilities.ArraySegment<T> cells) : base(shape_id, count)
        {
            this.Cells = cells;
        }
    }

    public class QueryOutputSections<T> : QueryOutputBase<T>
    {
        public List<SectionQueryOutput<T>> Sections { get; internal set; }

        internal QueryOutputSections(int shape_id, int count, List<SectionQueryOutput<T>> sections) : base(shape_id, count)
        {
            this.Sections = sections;
        }
    }
}