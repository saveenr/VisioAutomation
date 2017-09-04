using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Query
{
    public class QueryOutput<T> 
    {
        public int ShapeID { get; private set; }
        public VisioAutomation.Utilities.ArraySegment<T> Cells { get; internal set; }
        public List<SubQueryOutput<T>> Sections { get; internal set; }

        public int TotalCellCount;

        internal QueryOutput(int shape_id)
        {
            this.ShapeID = shape_id;
        }
    }

    public class QueryOutputCells<T>
    {
        public int ShapeID { get; private set; }
        public VisioAutomation.Utilities.ArraySegment<T> Cells { get; internal set; }

        public int TotalCellCount;

        internal QueryOutputCells(int shape_id)
        {
            this.ShapeID = shape_id;
        }
    }

    public class QueryOutputSections<T>
    {
        public int ShapeID { get; private set; }
        public List<SubQueryOutput<T>> Sections { get; internal set; }

        public int TotalCellCount;

        internal QueryOutputSections(int shape_id)
        {
            this.ShapeID = shape_id;
        }
    }
}