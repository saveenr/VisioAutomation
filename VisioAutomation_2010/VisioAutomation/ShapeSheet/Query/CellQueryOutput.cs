namespace VisioAutomation.ShapeSheet.Query
{
    public class CellQueryOutput<T>: QueryOutputBase<T>
    {
        public VisioAutomation.Utilities.ArraySegment<T> Cells { get; internal set; }

        internal CellQueryOutput(int shape_id, int count, VisioAutomation.Utilities.ArraySegment<T> cells) : base(shape_id, count)
        {
            this.Cells = cells;
        }
    }
}