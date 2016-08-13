namespace VisioAutomation.ShapeSheetQuery.Outputs
{
    public struct SubQueryOutputRow<T> 
    {
        public readonly T[] Cells;

        internal SubQueryOutputRow(T[] cells)
        {
            this.Cells = cells;
        }
    }
}