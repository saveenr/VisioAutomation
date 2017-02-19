using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Query
{
    public struct CellRange<T>
    {
        private readonly T[] array;

        private readonly T[] orig_array;
        private int StartIndex;
        private int length;

        public CellRange(T[] array, T[] orig_array, int startindex, int length)
        {
            this.array = array;
            this.orig_array = orig_array;
            this.StartIndex = startindex;
            this.length = length;
        }

        public T this[int index]
        {
            get
            {
                if (index >= this.length)
                {
                    throw new System.ArgumentOutOfRangeException(nameof(index));
                }

                T value1 = this.array[index];
                T value2 = this.orig_array[this.StartIndex + index];
               
                return value2;
            }
        }
    }
    public class QueryOutput<T> 
    {
        public int ShapeID { get; private set; }
        public CellRange<T> Cells { get; internal set; }
        public List<SubQueryOutput<T>> Sections { get; internal set; }

        //public int CursorStart;
        public int TotalCellCount;

        internal QueryOutput(int shape_id)
        {
            this.ShapeID = shape_id;
        }
    }
}