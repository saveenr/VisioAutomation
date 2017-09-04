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
}