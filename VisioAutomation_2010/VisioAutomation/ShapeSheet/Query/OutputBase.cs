namespace VisioAutomation.ShapeSheet.Query
{
    public class OutputBase 
    {
        public int ShapeID { get; private set; }

        internal readonly int __totalcellcount;

        internal OutputBase(int shapeid, int count)
        {
            this.ShapeID = shapeid;
            this.__totalcellcount = count;
        }
    }
}