namespace VisioAutomation.ShapeSheetQuery.Outputs
{
    public struct SubQueryOutputRow<T> 
    {
        public readonly T[] Cells;

        internal SubQueryOutputRow(T[] c)
        {
            this.Cells = c;
        }
    }
}