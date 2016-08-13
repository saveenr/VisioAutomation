using System.Collections.Generic;

namespace VisioAutomation.ShapeSheetQuery.Outputs
{
    public class Output<T> 
    {
        public int ShapeID { get; private set; }
        public T[] Cells { get; internal set; }
        public List<SubQueryOutput<T>> Sections { get; internal set; }

        internal Output(int sid)
        {
            this.ShapeID = sid;
        }
    }
}