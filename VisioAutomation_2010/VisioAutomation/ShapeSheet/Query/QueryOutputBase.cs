namespace VisioAutomation.ShapeSheet.Query
{
    public class QueryOutputBase<T> 
    {
        public int ShapeID { get; private set; }

        internal readonly int __totalcellcount;

        internal QueryOutputBase(int shapeid, int count)
        {
            this.ShapeID = shapeid;
            this.__totalcellcount = count;
        }
    }
}