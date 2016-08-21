using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Queries.Outputs
{
    public class Output<T> 
    {
        public int ShapeID { get; private set; }
        public T[] Cells { get; internal set; }
        public List<SubQueryOutput<T>> Sections { get; internal set; }
        internal int TotalCells;

        internal Output(int shape_id)
        {
            this.ShapeID = shape_id;
        }
    }
}