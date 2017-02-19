using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Query
{
    public struct CellRange<T>
    {
        private readonly T[] array;

        //private readonly T[] orig_array;
        //private int StartIndex;

        public CellRange(T[] array)
        {
            this.array = array;
            //this.orig_array = orig_array;
            //this.StartIndex = startpos;
        }

        public T this[int index]
        {
            get { return this.array[index]; }
            //set { /* set the specified index to value here */ }
        }
    }
    public class QueryOutput<T> 
    {
        public int ShapeID { get; private set; }
        public CellRange<T> Cells { get; internal set; }
        public List<SubQueryOutput<T>> Sections { get; internal set; }

        public int CursorStart;
        public int TotalCellCount;

        internal QueryOutput(int shape_id)
        {
            this.ShapeID = shape_id;
        }
    }
}