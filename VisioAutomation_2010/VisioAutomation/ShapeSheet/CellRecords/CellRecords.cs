namespace VisioAutomation.ShapeSheet.CellRecords
{
    public class CellRecords<T> : VisioAutomation.Core.BasicList<T> where T : CellRecord, new()
    {

        public CellRecords() : base()
        {

        }

        public CellRecords(int capacity) : base(capacity)
        {

        }

    }
}