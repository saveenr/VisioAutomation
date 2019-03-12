namespace VisioAutomation.ShapeSheet.Query
{
    public class CellQueryOutput<T>: QueryOutputBase<T>
    {
        public VisioAutomation.ShapeSheet.Internal.ArraySegment<T> Cells { get; internal set; }

        internal CellQueryOutput(int shape_id, int count, VisioAutomation.ShapeSheet.Internal.ArraySegment<T> cells) : base(shape_id, count)
        {
            this.Cells = cells;
        }
    }
}