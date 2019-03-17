namespace VisioAutomation.ShapeSheet.Query
{
    public class RowBase 
    {
        public int ShapeID { get; private set; }

        internal readonly int __totalcellcount;

        internal RowBase(int shapeid, int count)
        {
            this.ShapeID = shapeid;
            this.__totalcellcount = count;
        }
    }
}