namespace VisioAutomation.ShapeSheet.CellRecords
{
    public class CellRecordsGroup<T> : VisioAutomation.Core.BasicList<CellRecords<T>> where T : CellRecord, new()
    {
        public CellRecordsGroup(int capacity) : base(capacity)
        {

        }
    }
}